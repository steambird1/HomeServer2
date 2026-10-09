' print_excel.vbs
' Print xls / xlsx with MS Excel via COM
'
' Usage:
'   cscript //nologo print_excel.vbs <file> <printer> <orientation> <copies> <ranges> <firstOnly> [pdfOut] [maxWaitSec] [appearSec]
'     file        : full path of the workbook
'     printer     : printer name, "" = Excel default printer
'     orientation : 1 = xlPortrait, 2 = xlLandscape
'     copies      : number of copies (>=1)
'     ranges      : e.g. "1-3;5-5" (continuous ranges joined by ;), "" = all pages
'     firstOnly   : 1 = first worksheet only, 0 = every worksheet
'     pdfOut      : optional. if given, do NOT print: export to this PDF path and exit
'     maxWaitSec  : optional. longest time to wait for the job to finish (0 = no wait)
'     appearSec   : optional. grace period for the job to show up in the queue
'
' Exit code: 0 ok / 2 open failed / 3 print failed / 4 export failed / 9 bad arguments
'
' IMPORTANT (VBScript is NOT VBA):
'   1. Named arguments (Name:=value) are NOT supported - pass by position only.
'   2. Keep the file pure ASCII / ANSI with CRLF line endings.
'   3. Do not use the line-continuation "_" at end of line.
'
' HOW WE KNOW THE JOB REALLY FINISHED (primary = COM, fallback = print queue):
'   1. COM primary: Excel has no background-printing switch at all - PrintOut is
'      synchronous and only returns once the job has been handed to the spooler.
'   2. Fallback: track individual Win32_PrintJob JobIds plus the spooler's
'      cumulative TotalJobsPrinted counter (see lib_printq.vbs), because a job can
'      start and finish between two polls. Never judge by queue LENGTH alone, and
'      every wait has a hard upper bound.

Option Explicit

Dim args, sFile, sPrinter, nOrient, nCopies, sRanges, nFirstOnly, sPdfOut
Dim MAX_WAIT_SEC, APPEAR_SEC
Set args = WScript.Arguments

If args.Count < 6 Then
    WScript.Echo "ERROR: bad arguments (need 6, got " & args.Count & ")"
    WScript.Quit 9
End If

sFile = args(0)
sPrinter = args(1)
nOrient = CInt(args(2))
nCopies = CInt(args(3))
sRanges = args(4)
nFirstOnly = CInt(args(5))
sPdfOut = ""
MAX_WAIT_SEC = 300
APPEAR_SEC = 60
If args.Count >= 7 Then sPdfOut = args(6)
If args.Count >= 8 Then MAX_WAIT_SEC = CInt(args(7))
If args.Count >= 9 Then APPEAR_SEC = CInt(args(8))
If MAX_WAIT_SEC < 0 Then MAX_WAIT_SEC = 0
If APPEAR_SEC < 5 Then APPEAR_SEC = 5

If nCopies < 1 Then nCopies = 1

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

If Not fso.FileExists(sFile) Then
    WScript.Echo "ERROR: file not found: " & sFile
    WScript.Quit 2
End If

' ---- load the shared queue helpers (lib_printq.vbs, same folder) ----
Dim g_bQueueLib, sLibPath, oTs
g_bQueueLib = False
sLibPath = fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "lib_printq.vbs")
If fso.FileExists(sLibPath) Then
    On Error Resume Next
    Set oTs = fso.OpenTextFile(sLibPath, 1)
    If Err.Number = 0 Then
        ExecuteGlobal oTs.ReadAll
        oTs.Close
        ' ExecuteGlobal succeeded -> every PQ_* helper is now defined
        If Err.Number = 0 Then g_bQueueLib = True
    End If
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End If
If Not g_bQueueLib Then
    WScript.Echo "WARN: queue helper library not loaded - print completion cannot be verified"
End If

' Was Excel already running? Excel is a single-instance COM server: if the user
' has a workbook open, CreateObject attaches to THAT instance and Quit would close
' the user's window (and lose unsaved edits). Only Quit when we started Excel.
' Without the helper library we cannot tell, so we take the safe option: never Quit.
Dim bWeStartedExcel
If g_bQueueLib Then
    bWeStartedExcel = (PQ_ProcessCount("EXCEL.EXE") = 0)
Else
    bWeStartedExcel = False
End If

Dim xl, wb
On Error Resume Next
Set xl = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    WScript.Echo "ERROR: cannot start Excel: " & Err.Description
    WScript.Quit 2
End If
On Error GoTo 0

xl.Visible = False
xl.DisplayAlerts = False
xl.EnableEvents = False
xl.AskToUpdateLinks = False

If sPrinter <> "" Then
    On Error Resume Next
    xl.ActivePrinter = sPrinter
    If Err.Number <> 0 Then
        WScript.Echo "WARN: cannot select printer [" & sPrinter & "]: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
End If

' Workbooks.Open(FileName, UpdateLinks, ReadOnly)
On Error Resume Next
Set wb = xl.Workbooks.Open(sFile, 0, True)
If Err.Number <> 0 Or wb Is Nothing Then
    WScript.Echo "ERROR: cannot open workbook: " & Err.Description
    On Error Resume Next
    If bWeStartedExcel Then xl.Quit
    On Error GoTo 0
    WScript.Quit 2
End If
On Error GoTo 0

Dim idx, ws, arrRanges, i, pair, nFrom, nTo, nErrCount
Dim nErr, sErrDesc, sUsedPrinter
nErrCount = 0
If sRanges <> "" Then arrRanges = Split(sRanges, ";")
idx = 0

' Always pass a printer explicitly; if the caller gave none, use whatever Excel defaults to
If sPrinter = "" Then
    On Error Resume Next
    sUsedPrinter = xl.ActivePrinter
    If Err.Number <> 0 Then sUsedPrinter = ""
    On Error GoTo 0
Else
    sUsedPrinter = sPrinter
End If

' ---- snapshot the queue before printing so we can recognise our own job ----
Dim dictBefore, nPrintedBefore
Set dictBefore = CreateObject("Scripting.Dictionary")
If g_bQueueLib Then
    PQ_SnapshotJobs sPrinter, dictBefore
    nPrintedBefore = PQ_SpoolerTotalPrinted(sPrinter)
    WScript.Echo "INFO: queue before: " & dictBefore.Count & " job(s), spooler printed counter = " & nPrintedBefore
End If

For Each ws In wb.Worksheets
    idx = idx + 1
    If nFirstOnly = 1 And idx > 1 Then Exit For

    On Error Resume Next
    ws.PageSetup.Orientation = nOrient
    If Err.Number <> 0 Then
        WScript.Echo "WARN: cannot set orientation on sheet [" & ws.Name & "]: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    If sPdfOut <> "" Then
        ' export mode: orientation is already set, nothing to print yet
    ElseIf sRanges = "" Then
        ' PrintOut(From, To, Copies, Preview, ActivePrinter, PrintToFile, Collate)
        '   1 .. 9999 means "whole sheet" (Excel stops at the last real page)
        '   Arguments are passed by POSITION (VBScript has no named arguments)
        On Error Resume Next
        ws.PrintOut 1, 9999, nCopies, False, sUsedPrinter, False, True
        nErr = Err.Number
        sErrDesc = Err.Description
        On Error GoTo 0

        If nErr <> 0 Then
            WScript.Echo "WARN: PrintOut with arguments failed on sheet [" & ws.Name & "]: " & sErrDesc
            WScript.Echo "WARN: falling back to PrintOut without arguments"
            On Error Resume Next
            ws.PrintOut
            If Err.Number <> 0 Then
                WScript.Echo "WARN: fallback also failed on sheet [" & ws.Name & "]: " & Err.Description
                nErrCount = nErrCount + 1
            Else
                WScript.Echo "OK: sheet [" & ws.Name & "] sent to printer (fallback)"
            End If
            On Error GoTo 0
        Else
            WScript.Echo "OK: sheet [" & ws.Name & "] sent to printer"
        End If

    Else
        For i = 0 To UBound(arrRanges)
            pair = Split(arrRanges(i), "-")
            nFrom = CInt(pair(0))
            nTo = CInt(pair(1))

            On Error Resume Next
            ws.PrintOut nFrom, nTo, nCopies, False, sUsedPrinter, False, True
            nErr = Err.Number
            sErrDesc = Err.Description
            On Error GoTo 0

            If nErr <> 0 Then
                WScript.Echo "WARN: PrintOut with arguments failed on sheet [" & ws.Name & "]: " & sErrDesc
                WScript.Echo "WARN: falling back to PrintOut without arguments"
                On Error Resume Next
                ws.PrintOut
                If Err.Number <> 0 Then
                    WScript.Echo "WARN: fallback also failed on sheet [" & ws.Name & "]: " & Err.Description
                    nErrCount = nErrCount + 1
                Else
                    WScript.Echo "OK: sheet [" & ws.Name & "] sent to printer (fallback)"
                End If
                On Error GoTo 0
            Else
                WScript.Echo "OK: sheet [" & ws.Name & "] pages " & nFrom & "-" & nTo & " sent to printer"
            End If
        Next
    End If
Next

' ---- export-to-PDF mode (used by the duplex route) ----
' ExportAsFixedFormat(Type, Filename, ...)   0 = xlTypePDF
' NOTE: Excel takes Type FIRST and Filename SECOND (opposite of Word!)
If sPdfOut <> "" Then
    On Error Resume Next
    If nFirstOnly = 1 Then
        wb.Worksheets(1).ExportAsFixedFormat 0, sPdfOut
    Else
        wb.ExportAsFixedFormat 0, sPdfOut
    End If
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: export pdf failed: " & Err.Description
        On Error Resume Next
        wb.Close False
        If bWeStartedExcel Then xl.Quit
        On Error GoTo 0
        WScript.Quit 4
    End If
    On Error GoTo 0
    WScript.Echo "OK: exported pdf [" & sPdfOut & "]"
    On Error Resume Next
    wb.Close False
    If bWeStartedExcel Then xl.Quit
    On Error GoTo 0
    WScript.Quit 0
End If

' ---- confirm completion: COM already guaranteed the spool, queue scan double-checks ----
' Excel's PrintOut is synchronous, so the COM part is done by the time we get here.
Dim sRes, nCode, sNote
If nErrCount = 0 And MAX_WAIT_SEC > 0 And g_bQueueLib Then
    WScript.Echo "INFO: Excel PrintOut is synchronous - job already handed to the spooler"
    sRes = PQ_WaitForPrintDone(sPrinter, dictBefore, nPrintedBefore, MAX_WAIT_SEC, APPEAR_SEC)
    nCode = PQ_Code(sRes)
    sNote = PQ_Note(sRes)

    If nCode = 0 Then
        WScript.Echo "INFO: print confirmed by queue - " & sNote
    ElseIf nCode = 2 Then
        WScript.Echo "ERROR: print job failed - " & sNote
        nErrCount = nErrCount + 1
    Else
        WScript.Echo "WARN: UNCONFIRMED: " & sNote
    End If
End If

On Error Resume Next
wb.Close False
If bWeStartedExcel Then
    xl.Quit
    WScript.Echo "INFO: closed the Excel instance we started"
Else
    WScript.Echo "INFO: Excel was already running (or library missing) - left it open, only our workbook was closed"
End If
On Error GoTo 0

If nErrCount > 0 Then
    WScript.Echo "ERROR: " & nErrCount & " sheet/print operation(s) failed"
    WScript.Quit 3
End If

WScript.Quit 0

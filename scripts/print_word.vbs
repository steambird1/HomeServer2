' print_word.vbs
' Print doc / docx / txt with MS Word via COM
'
' Usage:
'   cscript //nologo print_word.vbs <file> <printer> <orientation> <copies> <pages> [pdfOut] [maxWaitSec] [appearSec]
'     file        : full path of the document
'     printer     : printer name, "" = Word default printer
'     orientation : 0 = portrait, 1 = landscape
'     copies      : number of copies (>=1)
'     pages       : e.g. "1-3,5,8-", "" = all pages
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
'   1. COM primary: PrintOut Background:=False (positional arg 1) blocks until the
'      job is spooled - same guarantee the Excel path has.
'   2. COM primary: poll Application.BackgroundPrintingStatus until it reaches 0.
'   3. Fallback: if that property is unavailable or never reaches 0, fall back to
'      tracking individual Win32_PrintJob JobIds plus the spooler's cumulative
'      TotalJobsPrinted counter (see lib_printq.vbs).
'   A job can start and finish between two polls, so we never rely on queue
'   LENGTH alone, and every wait has a hard upper bound.

Option Explicit

Dim args, sFile, sPrinter, nOrient, nCopies, sPages, sPdfOut
Dim MAX_WAIT_SEC, APPEAR_SEC
Set args = WScript.Arguments

If args.Count < 5 Then
    WScript.Echo "ERROR: bad arguments (need 5, got " & args.Count & ")"
    WScript.Quit 9
End If

sFile = args(0)
sPrinter = args(1)
nOrient = CInt(args(2))
nCopies = CInt(args(3))
sPages = args(4)
sPdfOut = ""
MAX_WAIT_SEC = 300
APPEAR_SEC = 60
If args.Count >= 6 Then sPdfOut = args(5)
If args.Count >= 7 Then MAX_WAIT_SEC = CInt(args(6))
If args.Count >= 8 Then APPEAR_SEC = CInt(args(7))
If MAX_WAIT_SEC < 0 Then MAX_WAIT_SEC = 0
If APPEAR_SEC < 5 Then APPEAR_SEC = 5

If nCopies < 1 Then nCopies = 1

Dim fso, sExt, nFormat, nEncoding
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

sExt = LCase(fso.GetExtensionName(sFile))

' txt: force format and encoding so the File Conversion dialog never pops up
nFormat = 0
nEncoding = 0
If sExt = "txt" Then
    nFormat = 7
    nEncoding = 65001
End If

' Was Word already running? Word is a single-instance COM server: if the user has
' a document open, CreateObject attaches to THAT instance and Quit would close
' the user's window (and lose unsaved edits). Only Quit when we started Word.
' Without the helper library we cannot tell, so we take the safe option: never Quit.
Dim bWeStartedWord
If g_bQueueLib Then
    bWeStartedWord = (PQ_ProcessCount("WINWORD.EXE") = 0)
Else
    bWeStartedWord = False
End If

Dim wd, doc
On Error Resume Next
Set wd = CreateObject("Word.Application")
If Err.Number <> 0 Then
    WScript.Echo "ERROR: cannot start Word: " & Err.Description
    WScript.Quit 2
End If
On Error GoTo 0

wd.Visible = False
wd.DisplayAlerts = 0
wd.ScreenUpdating = False

If sPrinter <> "" Then
    On Error Resume Next
    wd.ActivePrinter = sPrinter
    If Err.Number <> 0 Then
        WScript.Echo "WARN: cannot select printer [" & sPrinter & "]: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
End If

' Documents.Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles,
'                PasswordDocument, PasswordTemplate, Revert, WritePasswordDocument,
'                WritePasswordTemplate, Format, Encoding, Visible)
On Error Resume Next
If nFormat = 0 Then
    Set doc = wd.Documents.Open(sFile, False, True, False)
Else
    Set doc = wd.Documents.Open(sFile, False, True, False, "", "", False, "", "", nFormat, nEncoding, False)
End If

If Err.Number <> 0 Or doc Is Nothing Then
    WScript.Echo "ERROR: cannot open document: " & Err.Description
    On Error Resume Next
    If bWeStartedWord Then wd.Quit 0
    On Error GoTo 0
    WScript.Quit 2
End If
On Error GoTo 0

On Error Resume Next
doc.PageSetup.Orientation = nOrient
If Err.Number <> 0 Then
    WScript.Echo "WARN: cannot set orientation: " & Err.Description
    Err.Clear
End If
On Error GoTo 0

' ---- export-to-PDF mode (used by the duplex route) ----
' ExportAsFixedFormat(OutputFileName, ExportFormat, ...)  17 = wdExportFormatPDF
If sPdfOut <> "" Then
    On Error Resume Next
    doc.ExportAsFixedFormat sPdfOut, 17
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: export pdf failed: " & Err.Description
        On Error Resume Next
        doc.Close False
        If bWeStartedWord Then wd.Quit 0
        On Error GoTo 0
        WScript.Quit 4
    End If
    On Error GoTo 0
    WScript.Echo "OK: exported pdf [" & sPdfOut & "]"
    On Error Resume Next
    doc.Close False
    If bWeStartedWord Then wd.Quit 0
    On Error GoTo 0
    WScript.Quit 0
End If

' ---- snapshot the queue before printing so we can recognise our own job ----
Dim dictBefore, nPrintedBefore
Set dictBefore = CreateObject("Scripting.Dictionary")
If g_bQueueLib Then
    PQ_SnapshotJobs sPrinter, dictBefore
    nPrintedBefore = PQ_SpoolerTotalPrinted(sPrinter)
    WScript.Echo "INFO: queue before: " & dictBefore.Count & " job(s), spooler printed counter = " & nPrintedBefore
End If

' PrintOut(Background, Append, Range, OutputFileName, From, To, Item, Copies, Pages, ...)
'   Background = False  -> COM blocks until every page is spooled (PRIMARY guarantee)
'   Range 0 = wdPrintAllDocument, 4 = wdPrintRangeOfPages
'   Item  0 = wdPrintDocumentContent
'   Arguments are passed by POSITION (VBScript has no named arguments)
Dim nErr, sErrDesc, k
nErr = 0
sErrDesc = ""

If sPages <> "" Then
    On Error Resume Next
    doc.PrintOut False, False, 4, "", "", "", 0, nCopies, sPages
    nErr = Err.Number
    sErrDesc = Err.Description
    On Error GoTo 0
Else
    On Error Resume Next
    doc.PrintOut False, False, 0, "", "", "", 0, nCopies
    nErr = Err.Number
    sErrDesc = Err.Description
    On Error GoTo 0
End If

If nErr <> 0 Then
    WScript.Echo "WARN: PrintOut with arguments failed: " & sErrDesc
    WScript.Echo "WARN: falling back to PrintOut without arguments (page range ignored)"
    For k = 1 To nCopies
        On Error Resume Next
        doc.PrintOut
        nErr = Err.Number
        sErrDesc = Err.Description
        On Error GoTo 0
        If nErr <> 0 Then Exit For
    Next
End If

If nErr = 0 Then
    WScript.Echo "OK: document sent to printer [" & sPrinter & "] (COM synchronous PrintOut done)"
Else
    WScript.Echo "ERROR: print failed: " & sErrDesc
End If

' ---- confirm completion: COM first, queue scan only as fallback ----
Dim sRes, nCode, sNote, nConfirm, nShortWait
nConfirm = 0

If nErr = 0 And MAX_WAIT_SEC > 0 And g_bQueueLib Then
    ' primary: COM background counter (Word only)
    On Error Resume Next
    sRes = PQ_WaitWordBackground(wd, MAX_WAIT_SEC)
    If Err.Number <> 0 Then
        Err.Clear
        sRes = "3|BackgroundPrintingStatus not available on this Word build"
    End If
    On Error GoTo 0
    nCode = PQ_Code(sRes)
    sNote = PQ_Note(sRes)

    If nCode = 0 Then
        WScript.Echo "INFO: print confirmed by COM - " & sNote
        nConfirm = 1
    Else
        WScript.Echo "WARN: COM confirmation inconclusive - " & sNote
    End If

    ' fallback / cross-check: scan the queue by JobId
    If nConfirm = 1 Then
        nShortWait = 60
        If MAX_WAIT_SEC < nShortWait Then nShortWait = MAX_WAIT_SEC
    Else
        nShortWait = MAX_WAIT_SEC
    End If
    sRes = PQ_WaitForPrintDone(sPrinter, dictBefore, nPrintedBefore, nShortWait, APPEAR_SEC)
    nCode = PQ_Code(sRes)
    sNote = PQ_Note(sRes)

    If nCode = 0 Then
        WScript.Echo "INFO: print confirmed by queue - " & sNote
        nConfirm = 1
    ElseIf nCode = 2 Then
        WScript.Echo "ERROR: print job failed - " & sNote
        nErr = 1
        sErrDesc = sNote
    ElseIf nConfirm = 1 Then
        ' COM already confirmed; a missing job just means it finished instantly
        WScript.Echo "INFO: queue cross-check inconclusive (" & sNote & ") - COM already confirmed"
    Else
        WScript.Echo "WARN: UNCONFIRMED: " & sNote
    End If
End If

On Error Resume Next
doc.Close False
If bWeStartedWord Then
    wd.Quit 0
    WScript.Echo "INFO: closed the Word instance we started"
Else
    WScript.Echo "INFO: Word was already running (or library missing) - left it open, only our document was closed"
End If
On Error GoTo 0

If nErr <> 0 Then
    WScript.Quit 3
End If

WScript.Quit 0

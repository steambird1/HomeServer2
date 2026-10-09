' print_ppt.vbs
' Print ppt / pptx / pptm / pps / ppsx with MS PowerPoint via COM
'
' Usage:
'   cscript //nologo print_ppt.vbs <file> <printer> <orientation> <copies> <ranges> <perPage> <order> [pdfOut] [maxWaitSec] [appearSec]
'     file        : full path of the presentation
'     printer     : printer name, "" = current Windows default printer
'     orientation : 0 = portrait, 1 = landscape
'     copies      : number of copies (>=1)
'     ranges      : e.g. "1-3;5-5" (continuous slide ranges joined by ;), "" = every slide
'     perPage     : slides per sheet: 1, 2, 3, 4, 6 or 9  (handout mode when > 1)
'     order       : 0 = horizontal first, 1 = vertical first (handout mode only)
'     pdfOut      : optional. if given, do NOT print: export to this PDF path and exit
'     maxWaitSec  : optional. longest time to wait for the print job to complete
'     appearSec   : optional. grace period for the job to show up in the queue
'     fastQuit    : optional. 1 = kill the PowerPoint we started instead of Quit()
'                   (Quit blocks 15s+ on some Office builds); 0 = graceful Quit
'
' Exit code: 0 ok / 2 open failed / 3 print failed / 4 export failed / 9 bad arguments
'
' IMPORTANT (VBScript is NOT VBA):
'   1. Named arguments (Name:=value) are NOT supported - pass by position only.
'   2. Keep the file pure ASCII / ANSI with CRLF line endings.
'   3. Do not use the line-continuation "_" at end of line.
'
' GOTCHAS THAT COST A LOT OF TIME (do not "simplify" these away):
'   A. PrintOptions has NO "PrintType" property. Assigning it raises
'      "Object doesn't support this property or method". Setting OutputType alone
'      is what switches PowerPoint into handout mode.
'   B. PrintOptions.PrintInBackground defaults to msoTrue -> PrintOut returns
'      immediately and the job is spooled in the background. Quitting PowerPoint
'      right after that KILLS the pending job: the script reports success and
'      nothing ever comes out of the printer. We force foreground printing
'      (PrintInBackground = 0) so PrintOut blocks until the job is spooled.
'   C. PowerPoint is a single-instance COM server: CreateObject attaches to the
'      instance the logged-on user already has open. Calling Quit on it closes
'      THE USER'S PowerPoint windows. We only Quit when we started PowerPoint
'      ourselves (no POWERPNT.EXE running before CreateObject).
'   D. On Error Resume Next does NOT clear Err automatically - every guarded
'      block must Err.Clear, otherwise a stale error masks later success/failure.
'   E. Waiting: we never decide "printed" from the queue LENGTH alone, because a
'      job can start and finish between two polls. We track individual JobIds,
'      and cross-check the spooler's cumulative TotalJobsPrinted counter.

Option Explicit

Dim args, sFile, sPrinter, nOrient, nCopies, sRanges, nPerPage, nOrder, sPdfOut
Dim sRes, nCode, sNote, bFastQuit
Set args = WScript.Arguments

If args.Count < 7 Then
    WScript.Echo "ERROR: bad arguments (need 7, got " & args.Count & ")"
    WScript.Quit 9
End If

sFile = args(0)
sPrinter = args(1)
nOrient = CInt(args(2))
nCopies = CInt(args(3))
sRanges = args(4)
nPerPage = CInt(args(5))
nOrder = CInt(args(6))
sPdfOut = ""
If args.Count >= 8 Then sPdfOut = args(7)

Dim MAX_WAIT_SEC, APPEAR_SEC
MAX_WAIT_SEC = 300
APPEAR_SEC = 45
If args.Count >= 9 Then MAX_WAIT_SEC = CInt(args(8))
If args.Count >= 10 Then APPEAR_SEC = CInt(args(9))
bFastQuit = True
If args.Count >= 11 Then bFastQuit = (CInt(args(10)) = 1)
If MAX_WAIT_SEC < 10 Then MAX_WAIT_SEC = 10
If APPEAR_SEC < 5 Then APPEAR_SEC = 5

If nCopies < 1 Then nCopies = 1

' PowerPoint only accepts these slide counts per sheet
Dim bSupported
bSupported = False
If nPerPage = 1 Then bSupported = True
If nPerPage = 2 Then bSupported = True
If nPerPage = 3 Then bSupported = True
If nPerPage = 4 Then bSupported = True
If nPerPage = 6 Then bSupported = True
If nPerPage = 9 Then bSupported = True
If Not bSupported Then
    WScript.Echo "WARN: unsupported slides per page (" & nPerPage & "), using 1"
    nPerPage = 1
End If

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

' Was PowerPoint already running? (see gotcha C)
' Without the helper library we cannot tell, so we take the safe option: never Quit.
Dim bWeStartedPpt
If g_bQueueLib Then
    bWeStartedPpt = (PQ_ProcessCount("POWERPNT.EXE") = 0)
Else
    bWeStartedPpt = False
End If

Dim ppt, pres
On Error Resume Next
Set ppt = CreateObject("PowerPoint.Application")
If Err.Number <> 0 Then
    WScript.Echo "ERROR: cannot start PowerPoint: " & Err.Description
    WScript.Quit 2
End If
On Error GoTo 0

' Presentations.Open(FileName, ReadOnly, Untitled, WithWindow)
'   ReadOnly = -1 (msoTrue), WithWindow = 0 (msoFalse) -> no visible window
On Error Resume Next
Set pres = ppt.Presentations.Open(sFile, -1, 0, 0)
If Err.Number <> 0 Or pres Is Nothing Then
    WScript.Echo "ERROR: cannot open presentation: " & Err.Description
    On Error Resume Next
    PptShutdown ppt, bWeStartedPpt, bFastQuit
    On Error GoTo 0
    WScript.Quit 2
End If
On Error GoTo 0

WScript.Echo "INFO: PowerPoint " & ppt.Version & ", " & pres.Slides.Count & " slides, perPage=" & nPerPage

' PageSetup.SlideOrientation: 1 = msoOrientationHorizontal (landscape), 2 = vertical (portrait)
Dim nSlideOrient
If nOrient = 1 Then
    nSlideOrient = 1
Else
    nSlideOrient = 2
End If

On Error Resume Next
pres.PageSetup.SlideOrientation = nSlideOrient
If Err.Number <> 0 Then
    WScript.Echo "WARN: cannot set slide orientation: " & Err.Description
    Err.Clear
End If
On Error GoTo 0

' ---------- print options ----------
' PrintOptions.OutputType : 1 = one slide per page, 2 = 2-up, 3 = 3-up,
'                           4 = 6-up, 8 = 4-up, 9 = 9-up   (see gotcha A)
' PrintOptions.HandoutOrder : 1 = horizontal first, 2 = vertical first
Dim nOutput, nHandoutOrder
If nPerPage = 1 Then
    nOutput = 1
ElseIf nPerPage = 2 Then
    nOutput = 2
ElseIf nPerPage = 3 Then
    nOutput = 3
ElseIf nPerPage = 4 Then
    nOutput = 8
ElseIf nPerPage = 6 Then
    nOutput = 4
Else
    nOutput = 9
End If

If nOrder = 1 Then
    nHandoutOrder = 2
Else
    nHandoutOrder = 1
End If

' gotcha B: foreground printing, otherwise the job dies with PowerPoint
On Error Resume Next
pres.PrintOptions.PrintInBackground = 0
If Err.Number <> 0 Then
    WScript.Echo "WARN: cannot force foreground printing: " & Err.Description
    Err.Clear
End If
On Error GoTo 0

On Error Resume Next
pres.PrintOptions.OutputType = nOutput
If Err.Number <> 0 Then
    WScript.Echo "ERROR: cannot set OutputType=" & nOutput & ": " & Err.Description
    Err.Clear
    On Error GoTo 0
    On Error Resume Next
    pres.Saved = True
    pres.Close
    On Error GoTo 0
    PptShutdown ppt, bWeStartedPpt, bFastQuit
    On Error GoTo 0
    WScript.Quit 3
End If
On Error GoTo 0

If nPerPage > 1 Then
    On Error Resume Next
    pres.PrintOptions.HandoutOrder = nHandoutOrder
    If Err.Number <> 0 Then
        WScript.Echo "WARN: cannot set HandoutOrder: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
End If

' read back: a handout layout that silently did not apply is worse than an error
Dim nReadBack
nReadBack = -1
On Error Resume Next
nReadBack = pres.PrintOptions.OutputType
If Err.Number <> 0 Then
    WScript.Echo "WARN: cannot read back OutputType: " & Err.Description
    Err.Clear
End If
On Error GoTo 0
WScript.Echo "INFO: OutputType readback = " & nReadBack & " (expected " & nOutput & ")"

If nPerPage > 1 And nReadBack <> nOutput Then
    WScript.Echo "ERROR: handout layout did NOT apply (OutputType is " & nReadBack & ", expected " & nOutput & ") - aborting instead of printing the wrong layout"
    On Error Resume Next
    pres.Saved = True
    pres.Close
    On Error GoTo 0
    PptShutdown ppt, bWeStartedPpt, bFastQuit
    On Error GoTo 0
    WScript.Quit 3
End If

If nPerPage > 1 Then
    WScript.Echo "INFO: handout mode confirmed, " & nPerPage & " slides per sheet"
End If

' ---- export-to-PDF mode ----
' PowerPoint's export is notoriously picky: on some builds
'   ExportAsFixedFormat(Path, 2, 2) raises "Type mismatch" (0x800A000D).
' So we try a ladder of strategies and keep the first one that actually produces
' a non-empty file. PowerPoint cannot export the handout layout to PDF, so
' pdfOut is only used with 1 slide per page (the caller enforces this).
If sPdfOut <> "" Then
    sRes = PptExportPdf(ppt, sFile, sPdfOut)
    nCode = PptCode(sRes)
    sNote = PptNote(sRes)
    If nCode <> 0 Then
        WScript.Echo "ERROR: export pdf failed: " & sNote
        On Error Resume Next
        pres.Saved = True
        pres.Close
        On Error GoTo 0
        PptShutdown ppt, bWeStartedPpt, bFastQuit
        On Error GoTo 0
        WScript.Quit 4
    End If
    WScript.Echo "OK: exported pdf [" & sPdfOut & "] via " & sNote
    On Error Resume Next
    pres.Saved = True
    pres.Close
    On Error GoTo 0
    PptShutdown ppt, bWeStartedPpt, bFastQuit
    On Error GoTo 0
    WScript.Quit 0
End If

' ---- switch the Windows default printer ----
' PowerPoint.Application has no ActivePrinter, so the only way to pick a printer
' is to make it the Windows default for the duration of this job.
Dim net, sOldDefault, sNowDefault, bSwitched
Set net = CreateObject("WScript.Network")
sOldDefault = PQ_CurrentDefaultPrinter()
bSwitched = False
sNowDefault = sOldDefault

If sPrinter = "" Then
    bSwitched = True
ElseIf LCase(Trim(sPrinter)) = LCase(Trim(sOldDefault)) Then
    bSwitched = True
    WScript.Echo "INFO: target printer is already the default one"
Else
    On Error Resume Next
    net.SetDefaultPrinter sPrinter
    If Err.Number <> 0 Then
        WScript.Echo "WARN: SetDefaultPrinter failed: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    sNowDefault = PQ_CurrentDefaultPrinter()
    If LCase(Trim(sNowDefault)) = LCase(Trim(sPrinter)) Then
        bSwitched = True
        WScript.Echo "INFO: default printer switched to [" & sPrinter & "]"
    End If
End If

If Not bSwitched Then
    WScript.Echo "ERROR: cannot switch the Windows default printer to [" & sPrinter & "]"
    WScript.Echo "ERROR: the default printer is still [" & sNowDefault & "]"
    WScript.Echo "ERROR: aborted - refusing to print on the wrong printer"
    On Error Resume Next
    pres.Saved = True
    pres.Close
    On Error GoTo 0
    PptShutdown ppt, bWeStartedPpt, bFastQuit
    On Error GoTo 0
    WScript.Quit 3
End If

' ---- snapshot the print queue BEFORE printing (see gotcha E) ----
Dim dictBefore, nPrintedBefore
Set dictBefore = CreateObject("Scripting.Dictionary")
If g_bQueueLib Then
    PQ_SnapshotJobs sPrinter, dictBefore
    nPrintedBefore = PQ_SpoolerTotalPrinted(sPrinter)
    WScript.Echo "INFO: queue before: " & dictBefore.Count & " job(s), spooler printed counter = " & nPrintedBefore
End If

' ---- print ----
' Presentation.PrintOut(From, To, PrintToFile, Copies, Collate)
'   From / To are slide numbers; Collate = -1 (msoTrue)
'   With PrintInBackground = 0 this call blocks until the job is spooled.
Dim nTotal, nFrom, nTo, arrRanges, i, pair, nErr, sErrDesc
nTotal = pres.Slides.Count
nErr = 0
sErrDesc = ""

If sRanges = "" Then
    On Error Resume Next
    pres.PrintOut 1, nTotal, "", nCopies, -1
    nErr = Err.Number
    sErrDesc = Err.Description
    On Error GoTo 0
    If nErr = 0 Then WScript.Echo "OK: slides 1-" & nTotal & " sent to printer"
Else
    arrRanges = Split(sRanges, ";")
    For i = 0 To UBound(arrRanges)
        pair = Split(arrRanges(i), "-")
        nFrom = CInt(pair(0))
        nTo = CInt(pair(1))
        If nFrom < 1 Then nFrom = 1
        If nTo > nTotal Then nTo = nTotal

        On Error Resume Next
        pres.PrintOut nFrom, nTo, "", nCopies, -1
        If Err.Number <> 0 Then
            nErr = Err.Number
            sErrDesc = Err.Description
        End If
        On Error GoTo 0
        If nErr = 0 Then WScript.Echo "OK: slides " & nFrom & "-" & nTo & " sent to printer"
    Next
End If

' ---- wait for the job we just created (see gotcha E) ----
' COM primary: PrintInBackground = 0 already made PrintOut block until spooled.
If nErr = 0 And g_bQueueLib Then
    sRes = PQ_WaitForPrintDone(sPrinter, dictBefore, nPrintedBefore, MAX_WAIT_SEC, APPEAR_SEC)
    nCode = PQ_Code(sRes)
    sNote = PQ_Note(sRes)
    If nCode = 0 Then
        WScript.Echo "INFO: print confirmed by queue - " & sNote
    ElseIf nCode = 2 Then
        WScript.Echo "ERROR: print job failed - " & sNote
        nErr = 1
        sErrDesc = sNote
    Else
        WScript.Echo "WARN: UNCONFIRMED: " & sNote
    End If
End If

' restore the previous default printer
If sPrinter <> "" And sOldDefault <> "" Then
    If LCase(Trim(sOldDefault)) <> LCase(Trim(sPrinter)) Then
        On Error Resume Next
        net.SetDefaultPrinter sOldDefault
        If Err.Number = 0 Then
            WScript.Echo "INFO: default printer restored to [" & sOldDefault & "]"
        Else
            WScript.Echo "WARN: cannot restore default printer: " & Err.Description
        End If
        On Error GoTo 0
    End If
End If

On Error Resume Next
pres.Saved = True
pres.Close
On Error GoTo 0
If bWeStartedPpt Then
    PptShutdown ppt, True, bFastQuit
Else
    WScript.Echo "INFO: PowerPoint was already running - left it open, only our document was closed"
End If

If nErr <> 0 Then
    WScript.Echo "ERROR: print failed: " & sErrDesc
    WScript.Quit 3
End If

WScript.Quit 0


' ============================================================================
' PDF export ladder
' ============================================================================

' Wait up to ~5s for the exported file to appear and become non-empty.
Function PptFileReady(oFso, sPath)
    Dim i, oFile, nSize
    PptFileReady = False
    For i = 1 To 10
        On Error Resume Next
        If oFso.FileExists(sPath) Then
            Set oFile = oFso.GetFile(sPath)
            If Err.Number = 0 Then
                nSize = CLng(oFile.Size)
                If nSize > 0 Then
                    PptFileReady = True
                    Err.Clear
                    On Error GoTo 0
                    Exit Function
                End If
            End If
        End If
        Err.Clear
        On Error GoTo 0
        WScript.Sleep 500
    Next
End Function


' Try one export call. nMode selects the call; returns "code|note".
'   1 = ExportAsFixedFormat(Path, 2, 2)                    PDF + print intent
'   2 = ExportAsFixedFormat(Path, 2, 1)                    PDF + screen intent
'   3 = ExportAsFixedFormat(Path, 2, 2, 0, 1, 1, 0)        explicit optional args
'   4 = SaveAs(Path, 32)                                   ppSaveAsPDF
'   5 = SaveAs(Path, 32, 0)                                ppSaveAsPDF, no font embedding
'
' Try order: SaveAs FIRST. Real-world finding (Office 16.0, zh-CN): every
' ExportAsFixedFormat variant raises 0x800A000D "Type mismatch" no matter how the
' file was opened, while SaveAs(Path, ppSaveAsPDF) works and produces correct
' output. ExportAsFixedFormat is kept as a fallback for builds where it works
' (it also allows more options), but it is no longer tried first.
Function PptTryExport(oPres, sPath, nMode, oFso)
    Dim nE, sD
    PptTryExport = "1|failed"
    nE = 0
    sD = ""

    ' remove any leftover file so FileReady() cannot be fooled by a previous try
    On Error Resume Next
    If oFso.FileExists(sPath) Then oFso.DeleteFile sPath, True
    Err.Clear
    On Error GoTo 0

    On Error Resume Next
    If nMode = 1 Then
        oPres.ExportAsFixedFormat sPath, 2, 2
    ElseIf nMode = 2 Then
        oPres.ExportAsFixedFormat sPath, 2, 1
    ElseIf nMode = 3 Then
        oPres.ExportAsFixedFormat sPath, 2, 2, 0, 1, 1, 0
    ElseIf nMode = 4 Then
        oPres.SaveAs sPath, 32
    Else
        oPres.SaveAs sPath, 32, 0
    End If
    nE = Err.Number
    sD = Err.Description
    Err.Clear
    On Error GoTo 0

    If nE <> 0 Then
        PptTryExport = "1|mode " & nMode & " raised 0x" & Hex(nE) & " " & sD
        Exit Function
    End If

    If PptFileReady(oFso, sPath) Then
        PptTryExport = "0|mode " & nMode
        Exit Function
    End If

    PptTryExport = "1|mode " & nMode & " returned without error but produced no file"
End Function


' Export the open presentation to PDF, trying several COM calls.
' If the first pass fails we reopen the file WITH a visible window (some
' PowerPoint builds refuse to export from a windowless presentation) and, as a
' last resort, reopen it read-write (SaveAs needs a writable document).
' Returns "0|winning strategy" or "1|details".
Function PptExportPdf(pptApp, sSrcFile, sPdfOut)
    Dim oFso, sR, nC, i, arrModes

    Set oFso = CreateObject("Scripting.FileSystemObject")
    arrModes = Array(4, 5, 1, 2, 3)

    If oFso.FileExists(sPdfOut) Then
        On Error Resume Next
        oFso.DeleteFile sPdfOut, True
        Err.Clear
        On Error GoTo 0
    End If

    For i = 0 To UBound(arrModes)
        sR = PptTryExport(pres, sPdfOut, CLng(arrModes(i)), oFso)
        nC = PptCode(sR)
        If nC = 0 Then
            PptExportPdf = "0|" & PptModeName(CLng(arrModes(i))) & " (windowless)"
            Exit Function
        End If
        WScript.Echo "WARN: " & PptNote(sR)
    Next

    ' ---- pass 2: reopen WITH a window ----
    WScript.Echo "INFO: reopening the presentation with a window and retrying the export"
    On Error Resume Next
    pres.Saved = True
    pres.Close
    Err.Clear
    Set pres = pptApp.Presentations.Open(sSrcFile, -1, 0, -1)
    If Err.Number <> 0 Then
        WScript.Echo "WARN: cannot reopen with a window: " & Err.Description
        Err.Clear
        On Error GoTo 0
        PptExportPdf = "1|all strategies failed"
        Exit Function
    End If
    On Error GoTo 0

    For i = 0 To UBound(arrModes)
        sR = PptTryExport(pres, sPdfOut, CLng(arrModes(i)), oFso)
        nC = PptCode(sR)
        If nC = 0 Then
            PptExportPdf = "0|" & PptModeName(CLng(arrModes(i))) & " (windowed)"
            Exit Function
        End If
        WScript.Echo "WARN: " & PptNote(sR)
    Next

    ' ---- pass 3: reopen read-write (SaveAs may need it) ----
    WScript.Echo "INFO: reopening read-write and retrying SaveAs"
    On Error Resume Next
    pres.Saved = True
    pres.Close
    Err.Clear
    Set pres = pptApp.Presentations.Open(sSrcFile, 0, 0, 0)
    If Err.Number <> 0 Then
        WScript.Echo "WARN: cannot reopen read-write: " & Err.Description
        Err.Clear
        On Error GoTo 0
        PptExportPdf = "1|all strategies failed"
        Exit Function
    End If
    On Error GoTo 0

    For i = 0 To 1
        sR = PptTryExport(pres, sPdfOut, CLng(arrModes(i)), oFso)
        nC = PptCode(sR)
        If nC = 0 Then
            PptExportPdf = "0|" & PptModeName(CLng(arrModes(i))) & " (read-write)"
            Exit Function
        End If
        WScript.Echo "WARN: " & PptNote(sR)
    Next

    PptExportPdf = "1|all 5 COM calls failed on all 3 open modes; see WARN lines above"
End Function


' Shut down PowerPoint - but only the instance WE started (see gotcha C).
' bFast: PowerPoint.Quit() blocks for 15+ seconds on some Office 16 builds,
'        which shows up as a mysterious delay after "exported pdf" / "slides sent
'        to printer". Killing the process we started is instant and safe here,
'        because the presentation is already closed and the print job (if any) is
'        already spooled to Windows - it no longer depends on PowerPoint.
Sub PptShutdown(pptApp, bWeStarted, bFast)
    If Not bWeStarted Then Exit Sub
    On Error Resume Next
    If bFast Then
        Dim oSh
        Set oSh = CreateObject("WScript.Shell")
        If Err.Number = 0 Then
            ' 0 = hide window, False = do not wait for it
            oSh.Run "cmd /c taskkill /F /IM POWERPNT.EXE", 0, False
        End If
        If Err.Number <> 0 Then
            WScript.Echo "WARN: fast quit unavailable: " & Err.Description
            Err.Clear
            pptApp.Quit
        Else
            WScript.Echo "INFO: closed the PowerPoint instance we started (fast)"
        End If
    Else
        pptApp.Quit
        WScript.Echo "INFO: closed the PowerPoint instance we started"
    End If
    On Error GoTo 0
End Sub


Function PptCode(sRes)
    Dim nPos
    nPos = InStr(sRes, "|")
    If nPos > 1 Then
        PptCode = CLng(Left(sRes, nPos - 1))
    Else
        PptCode = 1
    End If
End Function


Function PptNote(sRes)
    Dim nPos
    nPos = InStr(sRes, "|")
    If nPos > 0 Then
        PptNote = Mid(sRes, nPos + 1)
    Else
        PptNote = sRes
    End If
End Function


Function PptModeName(nMode)
    If nMode = 1 Then
        PptModeName = "ExportAsFixedFormat(Path,2,2)"
    ElseIf nMode = 2 Then
        PptModeName = "ExportAsFixedFormat(Path,2,1)"
    ElseIf nMode = 3 Then
        PptModeName = "ExportAsFixedFormat(Path,2,2,0,1,1,0)"
    ElseIf nMode = 4 Then
        PptModeName = "SaveAs(Path,32)"
    Else
        PptModeName = "SaveAs(Path,32,0)"
    End If
End Function

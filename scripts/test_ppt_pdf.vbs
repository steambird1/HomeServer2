' test_ppt_pdf.vbs
' Standalone diagnostic: try every PowerPoint-to-PDF COM call and report which
' one works on THIS machine. It does not print anything.
'
' Usage:
'   cscript //nologo scripts\test_ppt_pdf.vbs "D:\tmp\demo.pptx" "D:\tmp\out.pdf"
'
' If you omit the output path the script uses <script folder>\ppt_pdf_test.pdf.
'
' Exit code: 0 = at least one strategy worked, 4 = none worked.
'
' Same VBScript rules as the other scripts: no Name:=value, no "_" continuation,
' pure ASCII with CRLF.

Option Explicit

Dim args, sSrc, sOut
Set args = WScript.Arguments

If args.Count < 1 Then
    WScript.Echo "usage: cscript //nologo test_ppt_pdf.vbs <pptx> [out.pdf]"
    WScript.Quit 9
End If

sSrc = args(0)
If args.Count >= 2 Then
    sOut = args(1)
Else
    Dim fso0
    Set fso0 = CreateObject("Scripting.FileSystemObject")
    sOut = fso0.BuildPath(fso0.GetParentFolderName(WScript.ScriptFullName), "ppt_pdf_test.pdf")
End If

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

If Not fso.FileExists(sSrc) Then
    WScript.Echo "ERROR: source file not found: " & sSrc
    WScript.Quit 9
End If

WScript.Echo "============================================================"
WScript.Echo " PowerPoint -> PDF diagnostic"
WScript.Echo "============================================================"
WScript.Echo " source : " & sSrc
WScript.Echo " target : " & sOut
WScript.Echo ""

Dim ppt, pres, bStarted
bStarted = (CountProc("POWERPNT.EXE") = 0)

On Error Resume Next
Set ppt = CreateObject("PowerPoint.Application")
If Err.Number <> 0 Then
    WScript.Echo "ERROR: cannot start PowerPoint: " & Err.Description
    WScript.Quit 9
End If
On Error GoTo 0

WScript.Echo " PowerPoint version: " & ppt.Version
WScript.Echo ""

' mode 1 = windowless + read-only   (what the print server uses)
' mode 2 = windowed   + read-only
' mode 3 = windowless + read-write
Dim nOpenMode
For nOpenMode = 1 To 3
    Dim sModeName
    If nOpenMode = 1 Then
        sModeName = "windowless, read-only"
    ElseIf nOpenMode = 2 Then
        sModeName = "windowed, read-only"
    Else
        sModeName = "windowless, read-write"
    End If
    WScript.Echo "--- open mode " & nOpenMode & ": " & sModeName & " ---"

    On Error Resume Next
    If nOpenMode = 1 Then
        Set pres = ppt.Presentations.Open(sSrc, -1, 0, 0)
    ElseIf nOpenMode = 2 Then
        Set pres = ppt.Presentations.Open(sSrc, -1, 0, -1)
    Else
        Set pres = ppt.Presentations.Open(sSrc, 0, 0, 0)
    End If
    If Err.Number <> 0 Then
        WScript.Echo "  open FAILED: " & Err.Description
        Err.Clear
        On Error GoTo 0
    Else
        On Error GoTo 0
        WScript.Echo "  opened, slides = " & pres.Slides.Count
        Call TryAll(pres, sOut)
        On Error Resume Next
        pres.Saved = True
        pres.Close
        Err.Clear
        On Error GoTo 0
    End If
    WScript.Echo ""
Next

On Error Resume Next
If bStarted Then ppt.Quit
On Error GoTo 0

If fso.FileExists(sOut) Then
    WScript.Echo "RESULT: SUCCESS - " & sOut & " exists"
    WScript.Quit 0
End If

WScript.Echo "RESULT: FAILED - no strategy produced a PDF"
WScript.Quit 4


' ---------------------------------------------------------------------------
Sub TryAll(oPres, sPath)
    Dim arr, i, nE, sD, sName

    arr = Array(4, 5, 1, 2, 3)
    For i = 0 To UBound(arr)
        sName = ModeName(CLng(arr(i)))
        If fso.FileExists(sPath) Then
            On Error Resume Next
            fso.DeleteFile sPath, True
            Err.Clear
            On Error GoTo 0
        End If

        nE = 0
        sD = ""
        On Error Resume Next
        If CLng(arr(i)) = 1 Then
            oPres.ExportAsFixedFormat sPath, 2, 2
        ElseIf CLng(arr(i)) = 2 Then
            oPres.ExportAsFixedFormat sPath, 2, 1
        ElseIf CLng(arr(i)) = 3 Then
            oPres.ExportAsFixedFormat sPath, 2, 2, 0, 1, 1, 0
        ElseIf CLng(arr(i)) = 4 Then
            oPres.SaveAs sPath, 32
        Else
            oPres.SaveAs sPath, 32, 0
        End If
        nE = Err.Number
        sD = Err.Description
        Err.Clear
        On Error GoTo 0

        If nE <> 0 Then
            WScript.Echo "  [" & arr(i) & "] " & sName & " -> ERROR 0x" & Hex(nE) & " " & sD
        ElseIf FileReady(sPath) Then
            WScript.Echo "  [" & arr(i) & "] " & sName & " -> OK"
        Else
            WScript.Echo "  [" & arr(i) & "] " & sName & " -> no error but no file"
        End If
    Next
End Sub


Function FileReady(sPath)
    Dim i, oFile
    FileReady = False
    For i = 1 To 10
        On Error Resume Next
        If fso.FileExists(sPath) Then
            Set oFile = fso.GetFile(sPath)
            If Err.Number = 0 Then
                If CLng(oFile.Size) > 0 Then
                    FileReady = True
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


Function ModeName(nMode)
    If nMode = 1 Then
        ModeName = "ExportAsFixedFormat(Path,2,2)"
    ElseIf nMode = 2 Then
        ModeName = "ExportAsFixedFormat(Path,2,1)"
    ElseIf nMode = 3 Then
        ModeName = "ExportAsFixedFormat(Path,2,2,0,1,1,0)"
    ElseIf nMode = 4 Then
        ModeName = "SaveAs(Path,32)"
    Else
        ModeName = "SaveAs(Path,32,0)"
    End If
End Function


Function CountProc(sExe)
    Dim oWmi, oCol, oProc, n
    CountProc = 0
    n = 0
    On Error Resume Next
    Set oWmi = GetObject("winmgmts:\\.\root\cimv2")
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    Set oCol = oWmi.ExecQuery("Select Name From Win32_Process")
    For Each oProc In oCol
        If UCase(CStr(oProc.Name)) = UCase(sExe) Then n = n + 1
    Next
    On Error GoTo 0
    CountProc = n
End Function

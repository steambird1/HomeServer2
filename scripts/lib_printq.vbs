' lib_printq.vbs
' Shared print-queue helpers, loaded by the print_*.vbs scripts via ExecuteGlobal.
'
' This file is NOT run directly. Each script does:
'     ExecuteGlobal fso.OpenTextFile(<path to this file>, 1).ReadAll
' If the include fails for any reason the script still prints (COM is the primary
' guarantee) and only loses the queue-based confirmation.
'
' Every "wait" helper returns "code|note" instead of writing to a script-level
' variable, so the calling script needs no global declared before ExecuteGlobal
' runs (VBScript + Option Explicit cannot see globals created at runtime).
'
'   code 0 = confirmed done      1 = timed out (still queued)
'        2 = job reported error  3 = could not confirm (never seen / no counter)
'
' IMPORTANT (VBScript is NOT VBA):
'   1. No Name:=value named arguments anywhere.
'   2. Pure ASCII / ANSI with CRLF line endings - cscript reads this with the
'      system ANSI code page.
'   3. No "_" line continuation. Keep every statement on one line.

Function PQ_LibName()
    PQ_LibName = "lib_printq 1.0"
End Function


Function PQ_CurrentDefaultPrinter()
    Dim oWmi, oCol, oPrinter
    PQ_CurrentDefaultPrinter = ""
    On Error Resume Next
    Set oWmi = GetObject("winmgmts:\\.\root\cimv2")
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    Set oCol = oWmi.ExecQuery("Select Name From Win32_Printer Where Default = True")
    For Each oPrinter In oCol
        PQ_CurrentDefaultPrinter = oPrinter.Name
    Next
    On Error GoTo 0
End Function


Function PQ_ProcessCount(sExe)
    Dim oWmi, oCol, oProc, n
    PQ_ProcessCount = 0
    n = 0
    On Error Resume Next
    Set oWmi = GetObject("winmgmts:\\.\root\cimv2")
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    Set oCol = oWmi.ExecQuery("Select Name From Win32_Process")
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    For Each oProc In oCol
        If UCase(CStr(oProc.Name)) = UCase(sExe) Then n = n + 1
    Next
    On Error GoTo 0
    PQ_ProcessCount = n
End Function


' Fill oDict with  JobId -> "Document [JobStatus]"  for one printer.
' Win32_PrintJob.Name looks like "<printer name>,<job id>"; printer names may
' contain commas, so split on the LAST comma.
Sub PQ_SnapshotJobs(sPrinterName, oDict)
    Dim oWmi, oCol, oJob, sName, nPos, sJobPrinter
    oDict.RemoveAll
    On Error Resume Next
    Set oWmi = GetObject("winmgmts:\\.\root\cimv2")
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    Set oCol = oWmi.ExecQuery("Select JobId, Name, Document, JobStatus, Status, PagesPrinted, TotalPages From Win32_PrintJob")
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    For Each oJob In oCol
        sName = CStr(oJob.Name)
        sJobPrinter = sName
        nPos = InStrRev(sName, ",")
        If nPos > 1 Then sJobPrinter = Left(sName, nPos - 1)
        If sPrinterName = "" Or LCase(Trim(sJobPrinter)) = LCase(Trim(sPrinterName)) Then
            oDict(CStr(oJob.JobId)) = CStr(oJob.Document) & " [" & CStr(oJob.JobStatus) & "]"
        End If
    Next
    On Error GoTo 0
End Sub


' Cumulative "jobs printed" counter of the spooler for one printer.
' This is the only reliable way to notice a job that both appeared AND finished
' between two queue polls. Returns -1 when the perf class is unavailable.
Function PQ_SpoolerTotalPrinted(sPrinterName)
    Dim oWmi, oCol, oQ, sName
    PQ_SpoolerTotalPrinted = -1
    On Error Resume Next
    Set oWmi = GetObject("winmgmts:\\.\root\cimv2")
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    Set oCol = oWmi.ExecQuery("Select Name, TotalJobsPrinted From Win32_PerfRawData_Spooler_PrintQueue")
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    For Each oQ In oCol
        sName = CStr(oQ.Name)
        If sPrinterName = "" Or LCase(Trim(sName)) = LCase(Trim(sPrinterName)) Then
            PQ_SpoolerTotalPrinted = CLng(oQ.TotalJobsPrinted)
            Exit Function
        End If
    Next
    On Error GoTo 0
End Function


Function PQ_JobHasError(sStatus)
    Dim s
    s = UCase(CStr(sStatus))
    PQ_JobHasError = False
    If InStr(s, "ERROR") > 0 Then PQ_JobHasError = True
    If InStr(s, "DELETED") > 0 Then PQ_JobHasError = True
    If InStr(s, "DELETING") > 0 Then PQ_JobHasError = True
    If InStr(s, "OFFLINE") > 0 Then PQ_JobHasError = True
    If InStr(s, "PAPEROUT") > 0 Then PQ_JobHasError = True
    If InStr(s, "PAPER OUT") > 0 Then PQ_JobHasError = True
    If InStr(s, "INTERVENTION") > 0 Then PQ_JobHasError = True
    If InStr(s, "BLOCKED") > 0 Then PQ_JobHasError = True
    If InStr(s, "JAM") > 0 Then PQ_JobHasError = True
End Function


' Wait until the print jobs we created (JobIds that were not there in oBefore)
' leave the queue. Returns "code|note".
Function PQ_WaitForPrintDone(sPrinterName, oBefore, nPrintedBefore, nMaxWaitSec, nAppearSec)
    Dim dictNow, dictSeen, nElapsed, nStep, bSeen, bAllGone
    Dim arrKeys, i, sId, sStatus, sLine, nPrintedAfter

    PQ_WaitForPrintDone = "3|unknown"
    Set dictSeen = CreateObject("Scripting.Dictionary")
    Set dictNow = CreateObject("Scripting.Dictionary")
    bSeen = False
    nElapsed = 0
    nStep = 3

    Do While nElapsed <= nMaxWaitSec
        PQ_SnapshotJobs sPrinterName, dictNow

        If dictNow.Count > 0 Then
            arrKeys = dictNow.Keys
            For i = 0 To UBound(arrKeys)
                sId = CStr(arrKeys(i))
                If Not oBefore.Exists(sId) Then
                    bSeen = True
                    If Not dictSeen.Exists(sId) Then
                        dictSeen(sId) = dictNow.Item(sId)
                        WScript.Echo "INFO: new print job " & sId & " -> " & dictNow.Item(sId)
                    End If
                End If
            Next
        End If

        If dictSeen.Count > 0 Then
            arrKeys = dictSeen.Keys
            bAllGone = True
            sLine = ""
            For i = 0 To UBound(arrKeys)
                sId = CStr(arrKeys(i))
                If dictNow.Exists(sId) Then
                    bAllGone = False
                    sStatus = CStr(dictNow.Item(sId))
                    If sLine = "" Then
                        sLine = sId & ": " & sStatus
                    Else
                        sLine = sLine & " | " & sId & ": " & sStatus
                    End If
                    If PQ_JobHasError(sStatus) Then
                        PQ_WaitForPrintDone = "2|print job " & sId & " reported [" & sStatus & "]"
                        Exit Function
                    End If
                End If
            Next

            If bAllGone Then
                PQ_WaitForPrintDone = "0|tracked job(s) left the queue after " & nElapsed & "s"
                Exit Function
            End If

            If nElapsed Mod 15 = 0 Then
                WScript.Echo "INFO: waiting for print job(s), elapsed " & nElapsed & "s -> " & sLine
            End If
        End If

        nPrintedAfter = PQ_SpoolerTotalPrinted(sPrinterName)
        If nPrintedBefore >= 0 And nPrintedAfter >= 0 Then
            If nPrintedAfter > nPrintedBefore Then
                PQ_WaitForPrintDone = "0|spooler printed counter went " & nPrintedBefore & " -> " & nPrintedAfter & " after " & nElapsed & "s"
                Exit Function
            End If
        End If

        If (Not bSeen) And nElapsed >= nAppearSec Then
            PQ_WaitForPrintDone = "3|no print job appeared in the queue within " & nAppearSec & "s"
            Exit Function
        End If

        WScript.Sleep nStep * 1000
        nElapsed = nElapsed + nStep
    Loop

    PQ_WaitForPrintDone = "1|print job(s) still in the queue after " & nMaxWaitSec & "s"
End Function


' Word only: Application.BackgroundPrintingStatus is the COM-side counter of
' pending background print jobs. This is the PRIMARY check for Word; the queue
' scan is only a fallback. Returns "code|note".
Function PQ_WaitWordBackground(oApp, nMaxWaitSec)
    Dim nElapsed, nStep, nStat
    PQ_WaitWordBackground = "3|BackgroundPrintingStatus unavailable"
    nElapsed = 0
    nStep = 2

    Do While nElapsed <= nMaxWaitSec
        On Error Resume Next
        nStat = CLng(oApp.BackgroundPrintingStatus)
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            PQ_WaitWordBackground = "3|cannot read Word BackgroundPrintingStatus"
            Exit Function
        End If
        On Error GoTo 0

        If nStat <= 0 Then
            PQ_WaitWordBackground = "0|Word background print queue is empty after " & nElapsed & "s"
            Exit Function
        End If

        If nElapsed Mod 10 = 0 Then
            WScript.Echo "INFO: Word still has " & nStat & " background print job(s), elapsed " & nElapsed & "s"
        End If
        WScript.Sleep nStep * 1000
        nElapsed = nElapsed + nStep
    Loop

    PQ_WaitWordBackground = "1|Word still reports " & nStat & " background print job(s) after " & nMaxWaitSec & "s"
End Function


' split "code|note"
Function PQ_Code(sRes)
    Dim nPos
    nPos = InStr(sRes, "|")
    If nPos > 1 Then
        PQ_Code = CLng(Left(sRes, nPos - 1))
    Else
        PQ_Code = 3
    End If
End Function


Function PQ_Note(sRes)
    Dim nPos
    nPos = InStr(sRes, "|")
    If nPos > 0 Then
        PQ_Note = Mid(sRes, nPos + 1)
    Else
        PQ_Note = sRes
    End If
End Function

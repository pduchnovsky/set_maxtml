' ======================================================================================
' Script: set_maxtml_sync.vbs
' Description: Monitors and enforces HDR MaxTML (Nits) on a specific display index.
'              Parses multi-monitor output to ensure only the target device is synced.
'              It utilizes set_maxtml.exe, so it must be in the same directory.
'
' Usage: set_maxtml_sync.vbs /v:[Value] /s:[Seconds] /m:[Index]
' 
' Arguments:
'   /v:[Nits]    - The target brightness value to maintain (e.g., 993).
'   /s:[Seconds] - How often to check the monitor state in seconds (e.g., 5).
'   /m:[Index]   - The monitor ID reported by the .exe (e.g., 1).
'
' Example, in windows shortcut target use:
'   "C:\Path\set_maxtml_sync.vbs" /v:993 /s:5 /m:1
'
' Note: This script uses a temp file for output capture because WshShell.Run allows 
'       for completely hidden execution (window style 0), preventing CMD flashes 
'       during active gaming.
'
' Author: Peter Duchnovsky (https://github.com/pduchnovsky)
' ======================================================================================

Option Explicit

Dim WshShell, fso, strPath, args, targetNits, sleepTime, targetMonitor, lockFile, lockStream, wmi, colProcess, objProcess
Dim tempFile, currentOutput, needsUpdate, monName, lines, i

Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
strPath = fso.GetParentFolderName(WScript.ScriptFullName)
Set args = WScript.Arguments.Named

' --- 1. ARGUMENT VALIDATION ---
' Required: /v:[Nits] /s:[Seconds] /m:[MonitorIndex]
If Not args.Exists("v") Or Not args.Exists("s") Or Not args.Exists("m") Then
    MsgBox "Error: Required arguments missing." & vbCrLf & vbCrLf & _
           "Usage: /v:[Nits] /s:[Seconds] /m:[MonitorIndex]" & vbCrLf & _
           "Example: /v:993 /s:5 /m:1", 16, "HDR Sync Tool"
    WScript.Quit
End If

targetNits    = args.Item("v")
sleepTime     = CInt(args.Item("s")) * 1000
targetMonitor = args.Item("m")

' --- 2. FILE & LOCK DEFINITIONS ---
tempFile = WshShell.ExpandEnvironmentStrings("%TEMP%") & "\hdr_check_status.txt"
lockFile = WshShell.ExpandEnvironmentStrings("%TEMP%") & "\set_maxtml_m" & targetMonitor & ".lock"

' --- 3. SINGLETON ENFORCEMENT ---
' Kills previous instances to prevent multiple scripts fighting over the same monitor
On Error Resume Next
If fso.FileExists(lockFile) Then
    fso.DeleteFile lockFile, True
    If Err.Number <> 0 Then
        Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
        Set colProcess = wmi.ExecQuery("Select * from Win32_Process Where Name = 'wscript.exe' AND CommandLine LIKE '%" & WScript.ScriptName & "%'")
        For Each objProcess in colProcess
            ' Only kill if it's the same script file AND targeting the same monitor index
            If InStr(objProcess.CommandLine, WScript.ScriptFullName) > 0 And InStr(objProcess.CommandLine, "/m:" & targetMonitor) > 0 Then
                objProcess.Terminate()
            End If
        Next
        WScript.Sleep 500
    End If
End If
On Error GoTo 0

Set lockStream = fso.OpenTextFile(lockFile, 2, True) 

' --- 4. MONITORING LOOP ---
Do
    ' Redirect exe output to temp file
    WshShell.Run "cmd /c " & Chr(34) & Chr(34) & strPath & "\set_maxtml.exe" & Chr(34) & " > " & Chr(34) & tempFile & Chr(34) & Chr(34), 0, True
    
    needsUpdate = True
    monName = "Monitor " & targetMonitor
    
    If fso.FileExists(tempFile) Then
        On Error Resume Next
        currentOutput = fso.OpenTextFile(tempFile, 1).ReadAll
        
        ' Split output into lines to find the correct monitor index
        lines = Split(currentOutput, vbCrLf)
        For i = 0 To UBound(lines)
            ' Look for the line starting with the target monitor number (e.g., "1:")
            If Left(Trim(lines(i)), Len(targetMonitor) + 1) = targetMonitor & ":" Then
                
                ' Parse monitor name: "1: Odyssey G85SB (993 nits)" -> "Odyssey G85SB"
                If InStr(lines(i), "(") > 0 Then
                    monName = Trim(Split(Split(lines(i), ":")(1), "(")(0))
                End If

                ' Check if target nits already exist on this specific line
                If InStr(lines(i), targetNits & " nits") > 0 Then
                    needsUpdate = False
                End If
                Exit For
            End If
        Next
        On Error GoTo 0
    End If

    ' Apply update if the specific monitor index is not at target nits
    If needsUpdate Then
        ' Executes: set_maxtml.exe [Index] [Nits]
        WshShell.Run Chr(34) & strPath & "\set_maxtml.exe" & Chr(34) & " " & targetMonitor & " " & targetNits, 0, False
        WshShell.LogEvent 4, "set_maxtml: Corrected " & monName & " (Index " & targetMonitor & ") to " & targetNits
    End If

    WScript.Sleep sleepTime
Loop

lockStream.Close

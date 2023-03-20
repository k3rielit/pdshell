Option Explicit

Const ForReading = 1

Dim objShell, objShellApp, objFSO, strScriptDir, strFilePath
Set objShell = CreateObject("WScript.Shell")
Set objShellApp = CreateObject("Shell.Application")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Check if the script is running with administrator privileges
If Not IsAdmin() Then
    WScript.Echo "[WBS] Run as administrator next time."
    WScript.Quit
End If

' Get the path to the script's directory and the path to the config file
strScriptDir = objFSO.GetParentFolderName(WScript.ScriptFullName)
strFilePath = strScriptDir & "\config.txt"

WScript.Echo "[WBS] Directory: " & strScriptDir
WScript.Echo "[WBS] Config: " & strFilePath

' Check if the config file exists and read its contents if it does
If objFSO.FileExists(strFilePath) Then
    Dim objFile, strLine, arrSplitLine
    Set objFile = objFSO.OpenTextFile(strFilePath, ForReading)
    Do Until objFile.AtEndOfStream
        strLine = objFile.ReadLine
        strLine = Trim(strLine) ' Remove any leading or trailing spaces
        ' Check if the line is not empty and does not start with #
        If Len(strLine) > 0 And Left(strLine, 1) <> "#" Then
            arrSplitLine = Split(strLine, ";")
            WScript.Echo "-----------------------< " & arrSplitLine(0) & " >-----------------------"

            ' Switch command type
            Select Case arrSplitLine(0)
                Case "Run"
                    Call WBS_Run(arrSplitLine,False)

                Case "RunAndWait"
                    Call WBS_Run(arrSplitLine,True)

                Case "AutoInstall"
                    Call WBS_AutoInstall(arrSplitLine)

                Case "CreateIcon"
                    WScript.Echo "[WBS] CreateIcon: " & strLine

                Case Else
                    WScript.Echo "[WBS] Unknown command: " & strLine

            End Select

        End If
    Loop
    objFile.Close
Else
    WScript.Echo "[WBS] Config file not found."
    WScript.Quit
End If


' Check if the script is running with administrator privileges
Private Function IsAdmin()
    On Error Resume Next
    objShell.RegRead("HKEY_USERS\S-1-5-19\Environment\TEMP")
    If Err.number = 0 Then 
        IsAdmin = True
    Else
        IsAdmin = False
    End If
    Err.Clear
    On Error goto 0
End Function

' Function for converting relative to absolute paths
Private Function Pathfinder(strPath)
    On Error Resume Next
    Dim strAbsolutePath
    If Not objFSO.DriveExists(objFSO.GetDriveName(strPath)) Then
        strAbsolutePath = objFSO.GetAbsolutePathName(strPath)
        WScript.Echo "[Pathfinder] Relative: " & strPath & " > Absolute: " & strAbsolutePath
    Else
        strAbsolutePath = strPath
    End If
    Pathfinder = strAbsolutePath
End Function

' Run an executable if it exists
Private Function WBS_Run(arrParams, boolWaitOnReturn)
    On Error Resume Next
    Dim strAbsolutePath
    If UBound(arrParams)>=1 And Len(arrParams(1)) > 0 Then
        strAbsolutePath = Pathfinder(arrParams(1))
        If objFSO.FileExists(strAbsolutePath) Then
            WScript.Echo "[Run] Running: " & strAbsolutePath
            objShell.Run chr(34) & strAbsolutePath & chr(34), 1, boolWaitOnReturn
        Else
            WScript.Echo "[Run] Path not found: " & strAbsolutePath
        End If
    End If
End Function

' Checks whether the file exists, if not, runs the installer
Private Function WBS_AutoInstall(arrParams)
    On Error Resume Next
    Dim strAbsolutePathFile, strAbsolutePathInstaller
    If UBound(arrParams)>=2 And Len(arrParams(1)) > 0 Then
        strAbsolutePathFile = Pathfinder(arrParams(1))
        strAbsolutePathInstaller = Pathfinder(arrParams(2))
        If Not objFSO.FileExists(strAbsolutePathFile) Then
            WScript.Echo "[Install] Installing: " & strAbsolutePathInstaller
            objShell.Run chr(34) & strAbsolutePathInstaller & chr(34), 1, True
        Else
            WScript.Echo "[Install] Already installed: " & strAbsolutePathFile
        End If
    End If
End Function
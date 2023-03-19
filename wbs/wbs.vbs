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

            ' Switch command type
            Select Case arrSplitLine(0)
                Case "Run"
                    Call WBS_Run(arrSplitLine,False)

                Case "RunAndWait"
                    Call WBS_Run(arrSplitLine,False)

                Case "AutoInstall"
                    WScript.Echo "[WBS] AutoInstall: " & strLine

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
    if Err.number = 0 Then 
        IsAdmin = True
    else
        IsAdmin = False
    end if
    Err.Clear
    On Error goto 0
End Function

' Run an executable if it exists
Function WBS_Run(arrParams, boolWaitOnReturn)
    Dim intReturnCode, absolutePath
    On Error Resume Next
    If(UBound(arrParams)>=1 And Len(arrParams(1)) > 0) Then
        If(arrParams(1)) Then

        End If
        WScript.Echo "[Run] " & arrParams(0)
        ' Run the executable and wait for it to finish
        WBS_Run = objShell.Run(strExecutablePath, 1, boolWaitOnReturn)
        WScript.Echo "[Run] Return code: " & intReturnCode
    End If
End Function

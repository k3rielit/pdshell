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

WScript.Echo "WBS v0.3"
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

                Case "CreateShortcut"
                    Call WBS_CreateShortcut(arrSplitLine)
                Case "CreateIcon"
                    Call WBS_CreateShortcut(arrSplitLine)
                Case "CreateLink"
                    Call WBS_CreateShortcut(arrSplitLine)

                Case "ExecuteSql"
                    Call WBS_ExecuteSql(arrSplitLine)

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
    On Error goto 0
End Function

' Function for creating a directory for a given relative/absolute file path 
Private Sub AutoCreateDirectory(strPath)
    On Error Resume Next
    Dim strDirectoryPath, strAbsolutePath
    strAbsolutePath = Pathfinder(strPath)
    strDirectoryPath = objFSO.GetParentFolderName(strAbsolutePath)
    If Not objFSO.FolderExists(strDirectoryPath) Then
        objFSO.CreateFolder(strDirectoryPath)
    End If
    On Error goto 0
End Sub

' Run an executable if it exists
' Run;ExecutablePath
' Run;ExecutablePath;Arguments
' RunAndWait;ExecutablePath
' RunAndWait;ExecutablePath;Arguments
Private Function WBS_Run(arrParams, boolWaitOnReturn)
    On Error Resume Next
    Dim strAbsolutePath, strRunParam
    If UBound(arrParams)>=1 And Len(arrParams(1)) > 0 Then
        strAbsolutePath = Pathfinder(arrParams(1))
        If objFSO.FileExists(strAbsolutePath) Then
            ' Run;Executable;Arguments
            If UBound(arrParams)>=2 Then
                strRunParam = chr(34) & strAbsolutePath & chr(34) & " " & arrParams(2)
            ' Run;Executable
            Else
                strRunParam = chr(34) & strAbsolutePath & chr(34)
            End If
            WScript.Echo "[Run] Running: " & strRunParam
            objShell.Run strRunParam, 1, boolWaitOnReturn
        Else
            WScript.Echo "[Run] Path not found: " & strAbsolutePath
        End If
    End If
    On Error goto 0
End Function

' Checks whether the file exists, if not, runs the installer
' AutoInstall;FilePath;InstallerPath
Private Function WBS_AutoInstall(arrParams)
    On Error Resume Next
    Dim strAbsolutePathFile, strAbsolutePathInstaller
    If UBound(arrParams)>=2 And Len(arrParams(1)) > 0 And Len(arrParams(2)) > 0 Then
        strAbsolutePathFile = Pathfinder(arrParams(1))
        strAbsolutePathInstaller = Pathfinder(arrParams(2))
        If Not objFSO.FileExists(strAbsolutePathFile) Then
            WScript.Echo "[Install] Installing: " & strAbsolutePathInstaller
            objShell.Run chr(34) & strAbsolutePathInstaller & chr(34), 1, True
        Else
            WScript.Echo "[Install] Already installed: " & strAbsolutePathFile
        End If
    End If
    On Error goto 0
End Function

' Creates a shortcut / shell link (.lnk)
' CreateShortcut;ShortcutPath;TargetPath
' CreateIcon;ShortcutPath;TargetPath
' CreateLink;ShortcutPath;TargetPath
Private Function WBS_CreateShortcut(arrParams)
    On Error Resume Next
    Dim objShortcut, strShortcutPath, strTargetPath, strWorkingDirectoryPath
    If UBound(arrParams)>=2 And Len(arrParams(1)) > 0 And Len(arrParams(2)) > 0 Then
        strShortcutPath = Pathfinder(arrParams(1))
        strTargetPath = Pathfinder(arrParams(2))
        ' CreateShortcut() doesn't create the directory automatically
        strWorkingDirectoryPath = objFSO.GetParentFolderName(strTargetPath)
        AutoCreateDirectory strShortcutPath
        ' Create shortcut
        Set objShortcut = objShell.CreateShortcut(strShortcutPath)
        objShortcut.TargetPath = strTargetPath
        objShortcut.WorkingDirectory = strWorkingDirectoryPath
        objShortcut.WindowStyle = 1 ' 1 = Normal window
        objShortcut.IconLocation = strTargetPath & ",0"
        objShortcut.Save
        WScript.Echo "[CreateShortcut] Created successfully: " & strShortcutPath 
    End If
    On Error goto 0
End Function

' Executes a SQL command with the connection string
' Depends on ODBC Connector: https://dev.mysql.com/downloads/connector/odbc/
' ExecuteSql;driver;server,database,uid,pwd;SQL
Private Sub WBS_ExecuteSql(arrParams)
    On Error Resume Next
    Dim objConnection, objCommand
    If UBound(arrParams)>=5 Then
        ' Create a connection to the MySQL server
        Set objConnection = CreateObject("ADODB.Connection")
        objConnection.ConnectionString = "Driver=" & arrParams(1) & ";Server=" & arrParams(2) & ";Database=" & arrParams(3) & ";User=" & arrParams(4) & ";Password=" & arrParams(5) & ";"
        objConnection.Open
        WScript.Echo "[ExecuteSql] Connection state: " & objConnection.State
        ' Create a command object to execute the SQL statement
        Set objCommand = CreateObject("ADODB.Command")
        objCommand.ActiveConnection = objConnection
        objCommand.CommandText = arrParams(6) & ";"
        ' Execute the SQL statement
        WScript.Echo "[ExecuteSql] Executing: " & arrParams(6) & ";"
        objCommand.Execute
        If Err.Number <> 0 Then
            WScript.Echo "[ExecuteSql] Error: " & Err.Description
        End If
        ' Close the connection
        objConnection.Close
        Set objConnection = Nothing
        Set objCommand = Nothing
    End If
    On Error goto 0
End Sub
Option Explicit

Dim args, arg
Set args = Wscript.Arguments

If WScript.Arguments.Count = 0 Then
    WScript.Echo "Missing arguments"
Else
    For Each arg In args
      WScript.Echo arg
    Next
End If

WScript.Echo "Done, press any key to continue"
WScript.StdIn.Read(1)
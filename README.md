## PDShell
Simple self-elevating PowerShell automatization script.
## How To Use

* [PS <6.0] Set execution policy in an admin shell:

```powershell
Set-ExecutionPolicy Unrestricted -Force
# The policy affects all users by default.
```

* Edit config.txt with the following format:

```powershell
# Comment
C:\Absolute\path\executable.exe;.\relative\path\installer.exe;Icon Name
.\relative\path\executable2.exe;C:\Absolute\path\installer2.exe;2nd Icon Name
# ...
# It checks whether the executable exists, if not, runs the installer,
# and finally creates an icon for it in Public\Desktop.
```

* Right click on `pd.ps1` > `Run with PowerShell`
* [PS <6.0] Optional: Revert execution policy changes:

```powershell
Set-ExecutionPolicy Restricted -Force
Get-ExecutionPolicy -List
```

## Notes
It might crash if it's ran from a network drive. Probably related to the way relative paths are processed.
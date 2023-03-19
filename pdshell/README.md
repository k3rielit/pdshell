# PDShell

Really simple PowerShell script for checking whether a program is installed, if not, running an installer, then creating an icon on the public desktop for it. Relative paths ( `.\` ) are converted to absolute paths during runtime. Commented out lines are skipped ( starting with `#` ) while reading the config file.

## HOW TO USE

* [PS <6.0] Set execution policy in an admin shell:

```powershell
Set-ExecutionPolicy Unrestricted -Force # It affects all users by default.
Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope CurrentUser # For current user.
```

* Edit config.txt with the following format:

```powershell
# executable;installer;iconName
C:\Absolute\path\executable.exe;.\relative\path\installer.exe;Icon Name
.\relative\path\executable2.exe;C:\Absolute\path\installer2.exe;2nd Icon Name
```

* Right click on `pd.ps1` > `Run with PowerShell`
* [PS <6.0] Optional: Revert execution policy changes:

```powershell
Set-ExecutionPolicy Restricted -Force # For all users.
Get-ExecutionPolicy -List # See if it worked
```

## NOTES

The self-elevation might not work (right click run method). If the script closes itself, it failed to elevate. Opening up an admin shell and running it manually always works. Files hosted on a network drive need the real path, not the mounted path ( `\\192.168.0.200\Files\installer.exe` will work, `M:\installer.exe` won't ).

## TODO

* More features:
  * Copying files
  * Moving files
  * Deleting files
* Skipping empty lines

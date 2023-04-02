# WBS

VBScript for Windows automatization via a config file.

## HOW TO USE

* Create a `config.txt` file next to `wbs.vbs`.
* Add comment, empty, or command lines to it.
* `CLI` Run `wbs.bat` as administrator.

### CONFIG FILE

Comments, empty lines:

```ps
# Comment 1

# Comment 2
```

Command lines:

```ps
CommandType;param1
CommandType;param1;param2
CommandType;param1;param2;param3
```

```ps
Run;path\program.exe                   # Run the executable, WaitOnReturn = False
Run;path\program.exe;Arguments         # Run the executable, WaitOnReturn = False
RunAndWait;path\program.exe;Arguments  # Run the executable, WaitOnReturn = True
RunAndWait;path\program.exe            # Run the executable, WaitOnReturn = True
AutoInstall;path\to\file;path\to\installer.exe  # If the file doesn't exist, runs the installer
CreateShortcut;shortcut\path\Shortcut.lnk;target\file.txt   # Creates a shortcut
CreateIcon;icon\path\Icon.lnk;target\directory\             # Creates a shortcut
CreateLink;icon\path\Icon.lnk;target\executable\program.exe # Creates a shortcut
```

## NOTES

Double clicking `wbs.vbs` runs it via `wscript`, where outputs are using message boxes (Windows application). Run it via `cscript` instead to have a CLI (console application). ( [Source](https://stackoverflow.com/a/9062764), [Source](http://scripts.dragon-it.co.uk/scripts.nsf/MainFrame?OpenFrameSet&Frame=East&Src=%2Fscripts.nsf%2Fdocs%2Fvbscript-writing-to-stdout-stderr!OpenDocument%26AutoFramed) )

The script requires administrator privileges, or else it'll exit. But CMD or PowerShell windows started as admin always has System32 as their working directory. The `cli.bat` solves this issue. ( [Source](https://stackoverflow.com/a/30256894) ) But it won't work if ran from a network path (like `\\192.168.0.200\path\`), because CMD doesn't support UNC paths as current directories.

To check whether the script was run as an administrator, trying to access a protected registry key is the simplest way. ( [Source](https://stackoverflow.com/a/45069476) )

Paths can be:

* Relative: `.\relative\path\`
* Absolute: `C:\absolute\path\`, `\\192.168.0.200\network\path\`

The `CreateShortcut`/`CreateIcon`/`CreateLink` commands work for executables, regular files, and directories as well. They're' just synonyms for the same command.

## TODO

* Workaround for CMD not supporting UNC paths (cli.bat) ( [Source](https://superuser.com/questions/282963/browse-an-unc-path-using-windows-cmd-without-mapping-it-to-a-network-drive) )

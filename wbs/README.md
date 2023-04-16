# WBS

VBScript for Windows automatization via a config file.

## HOW TO USE

* [Option 1] Edit cli.bat:
  * `cscript "<wbs.vbs absolute path>" "command;arg" "command;arg"`
* [Option 2] Create a shell link (.lnk):
  * Target: `C:\Windows\System32\cscript.exe "C:\path\to\wbs.vbs" "command;arg" "command;arg"`
* Use `SetRootPath;path` if you're using relative paths, because it'll be System32

### CONFIG FILE

Comments, empty lines:

```ps
# Comment 1

# Comment 2
```

Command lines:

```ps
Command;param1
Command;param1;param2
Command;param1;param2;param3
```

```ps
ProcessConfig;path\config.txt # Reads the file and executes commands
SetRootPath;Path # Changes script root to custom path in the relative > absolute path converter
UnsetRootPath    # Resets root path to default (script root)
DefaultRootPath  # Alias of UnsetRootPath, resets root path to default (script root)
PressAnyKey         # Waits for keypress
PressAnyKey;Message # Waits for keypress, displays the specified message
Run;path\program.exe                   # Run the executable, WaitOnReturn = False
Run;path\program.exe;Arguments         # Run the executable, WaitOnReturn = False
RunAndWait;path\program.exe;Arguments  # Run the executable, WaitOnReturn = True
RunAndWait;path\program.exe            # Run the executable, WaitOnReturn = True
AutoInstall;path\to\file;path\to\installer.exe      # If the file doesn't exist, runs the installer
AutoInstall;path\to\file;path\to\installer.exe;args # If the file doesn't exist, runs the installer
CreateShortcut;shortcut\path\Shortcut.lnk;target\file.txt   # Creates a shortcut
CreateIcon;icon\path\Icon.lnk;target\directory\             # Creates a shortcut
CreateLink;icon\path\Icon.lnk;target\executable\program.exe # Creates a shortcut
ExecuteSql;driver;server,database,uid,pwd;SQL               # Executes SQL
Uninstall;DisplayName # Searches in the registry for *name* and uninstalls every occurrence
```

## NOTES

Double clicking `wbs.vbs` runs it via `wscript`, where outputs are using message boxes (Windows application). Run it via `cscript` instead to have a CLI (console application). ( [Source](https://stackoverflow.com/a/9062764), [Source](http://scripts.dragon-it.co.uk/scripts.nsf/MainFrame?OpenFrameSet&Frame=East&Src=%2Fscripts.nsf%2Fdocs%2Fvbscript-writing-to-stdout-stderr!OpenDocument%26AutoFramed) )

The script requires administrator privileges, or else it'll exit. But CMD or PowerShell windows started as admin always has System32 as their working directory. The `cli.bat` solves this issue. ( [Source](https://stackoverflow.com/a/30256894) ) But it won't work if ran from a network path (like `\\192.168.0.200\path\`), because CMD doesn't support UNC paths as current directories.

To check whether the script was run as an administrator, trying to access a protected registry key is the simplest way. ( [Source](https://stackoverflow.com/a/45069476) )

Paths can be:

* Relative: `.\relative\path\`
* Absolute: `C:\absolute\path\`, `\\192.168.0.200\network\path\`

The `CreateShortcut`/`CreateIcon`/`CreateLink` commands work for executables, regular files, and directories as well. They're' just synonyms for the same command.

The `ExecuteSql` command depends on the ODBC Connector ( [Download](https://dev.mysql.com/downloads/connector/odbc/) ). It needs the driver to be specified, it'll be `{MySQL ODBC 8.0 Unicode Driver}` for MySQL 8.0 for example.

The `Uninstall` command searches for a program's key in *HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall*, then uninstalls it using MsiExec's passive mode (only a progressbar pops up).

## TODO

* Workaround for CMD not supporting UNC paths (cli.bat) ( [Source](https://superuser.com/questions/282963/browse-an-unc-path-using-windows-cmd-without-mapping-it-to-a-network-drive) )
* Use shell links instead of cli.bat
* Commands as arguments
* Move command processor into Sub
* Dynamically link config files instead of the hardcoded config.txt
* Better documentation for commands

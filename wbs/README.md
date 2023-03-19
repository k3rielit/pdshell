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
Run;path\program.exe # Run the executable, WaitOnReturn = False
RunAndWait;path\program.exe # Run the executable, WaitOnReturn = True
```

## NOTES

Double clicking `wbs.vbs` runs it via `wscript`, where outputs are using message boxes (Windows application). Run it via `cscript` instead to have a CLI (console application). ( [Source](https://stackoverflow.com/a/9062764), [Source](http://scripts.dragon-it.co.uk/scripts.nsf/MainFrame?OpenFrameSet&Frame=East&Src=%2Fscripts.nsf%2Fdocs%2Fvbscript-writing-to-stdout-stderr!OpenDocument%26AutoFramed) )

The script requires administrator privileges, or else it'll exit. But CMD or PowerShell windows started as admin always has System32 as their working directory. The `cli.bat` solves this issue. ( [Source](https://stackoverflow.com/a/30256894) )

To check whether the script was run as an administrator, trying to access a protected registry key is the simplest way. ( [Source](https://stackoverflow.com/a/45069476) )

## TODO

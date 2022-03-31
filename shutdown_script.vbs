Option Explicit
Dim objShell, Computer, Shutdown

Computer = "NWXFS01.nwxdemo.lcl"

Shutdown = "shutdown -s -t 0 -f -m \\" & Computer

set objShell = CreateObject("WScript.Shell")

objShell.run Shutdown

Wscript.Quit
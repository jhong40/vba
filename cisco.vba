Dim objShell
set objShell = WScript.CreateObject("WScript.Shell")

objShell.run """C:\Program Files (x86)\Cisco\Cisco AnyConnect Secure Mobility Client\vpnui.exe"""
WScript.Sleep  1000
objShell.SendKeys "{ENTER}"
WScript.Sleep  1000
objShell.SendKeys "mypass"
objShell.SendKeys "{TAB}"

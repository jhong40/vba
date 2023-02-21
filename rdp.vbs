Dim objShell, strMachineName, strUserName, strUserPwd
set objShell = WScript.CreateObject("WScript.Shell")
strMachineName = "1.2.3.4"
'strUserName = "domain\username"
'strUserPwd = "mypassword"
'objShell.Run "cmdkey /generic:"&strMachineName&" /user:"&strUserName&" /pass:"&strUserPwd
objShell.run "mstsc /v: "&strMachineName

WScript.Sleep  1000

objShell.SendKeys "mypasswordblah"
objShell.SendKeys "{ENTER}"

set objShell = Nothing

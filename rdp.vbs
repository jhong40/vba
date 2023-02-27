Set fso = CreateObject("Scripting.FileSystemObject")   
Set bFile = fso.OpenTextFile("note.txt", 1) 
myStr = bFile.ReadLine
bFile.Close                       
Set bFile = Nothing 
'MsgBox myStr


Dim objShell, strMachineName, strUserName, strUserPwd
set objShell = WScript.CreateObject("WScript.Shell")
strMachineName = "1.2.3.4"

objShell.run "mstsc /v: "&strMachineName

WScript.Sleep  1000*2


objShell.SendKeys myStr
objShell.SendKeys "{ENTER}"

set objShell = Nothing

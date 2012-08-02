' Запустить Net-Worm.Win32.Kido removing tool, Kaspersky Lab 2010
' отправить логи на почту
' im.vbs /antivir_kido
'
' @file       kido.vbs
' @copyright  Copyright (c) 2011 iMega ltd. (http://www.imega.ru, info@imega.ru)
' @version    0.1

includeClass ("iMega_Shell")
dim shell: set shell = new iMega_Shell

dim objNetwork: set objNetwork = CreateObject("WScript.Network")
dim strComputerName: strComputerName = objNetwork.ComputerName

dim resultShell: resultShell = shell.cmd("\\alliance.local\NETLOGON\soft\kk.exe -l c:\kk.log -y")
'if resultShell = 0 then
	set fso = CreateObject("Scripting.FileSystemObject")
	set InpFile = fso.OpenTextFile("c:\kk.log", 1, False, -1)
	InpStr=InpFile.ReadAll
'else
'	InpStr = "Error"
'	strComputerName = "ОШИБКА " & strComputerName
'end if 

includeClass ("iMega_Mail")
dim oMail: set oMail = new iMega_Mail
with oMail
	.from = "kido.remover@alliance-motors.ru"
	.message = InpStr	
	.recipient = "sysadmin@alliance-motors.ru"
	.server = "smtp.alliance-motors.ru"
	.subject = strComputerName & " Net-Worm.Win32.Kido removing tool, Kaspersky Lab 2010"
	.send
end with
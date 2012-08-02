' Установка Microsoft Office 14 (2010)
' запуск: im.vbs /install_office_install
'
' @file       install.vbs
' @copyright  Copyright (c) 2011 iMega ltd. (http://www.imega.ru, info@imega.ru)
' @version    0.1

dim sNameProg: sNameProg = "Microsoft Office, для*"

' Если Workstation то продолжаем
includeClass ("iMega_OS"): dim oOs: set oOs = new iMega_OS

if oOs.productType > 1 then quit(0)

includeClass ("iMega_FS"): dim oFS: set oFS = new iMega_FS
includeClass ("iMega_PC"): dim oPC: set oPC = new iMega_PC

dim pcName: pcName = oPC.name

includeClass ("iMega_Program"): dim oProg: set oProg = new iMega_Program

dim iVersion: iVersion = ""

with oProg
	.search sNameProg
	iVersion = .displayName
end with

'echo (iVersion)

dim mspFile: mspFile = "\\alliance.local\netlogon\soft\office-msp\" & pcName & ".msp"
dim resultExists: resultExists = oFS.existsFile(mspFile)
if iVersion = "" and resultExists then
	includeClass ("iMega_Shell")
	dim oShell: set oShell = new iMega_Shell
	dim resultShell: resultShell = oShell.cmd("\\data\install\editors\text\office14lic\x" & oOs.osArchitecture & "\setup.exe /adminfile " & mspFile)
	if resultShell = 0 then
		quit(0)
	else
		quit(1)
	end if
end if
' Установка Microsoft Security Essentials
' запуск: im.vbs /install_mse
'
' @file       mse.vbs
' @copyright  Copyright (c) 2011 iMega ltd. (http://www.imega.ru, info@imega.ru)
' @version    0.1

dim sNameProg: sNameProg = "Microsoft Security Essentia*"

' Если Workstation то продолжаем
includeClass ("iMega_OS"): dim oOs: set oOs = new iMega_OS
if oOs.productType > 1 then quit(0)

includeClass ("iMega_FS"): dim oFS: set oFS = new iMega_FS
includeClass ("iMega_PC"): dim oPC: set oPC = new iMega_PC

dim pcName: pcName = oPC.name

includeClass ("iMega_Program"): dim oProg: set oProg = new iMega_Program

dim iVersion: iVersion = ""
dim nod: nod = ""
with oProg
	.search "ESET NOD*"
	nod = .displayName
	.search sNameProg
	iVersion = .displayVersion
end with

if iVersion = "" and nod = "" then
	includeClass ("iMega_Shell")
	dim oShell: set oShell = new iMega_Shell
	dim resultShell: resultShell = oShell.cmd("\\data\install\antivirus\mse\x" & oOs.osArchitecture & "\mseinstall.exe /s /runwgacheck /o")
	if resultShell = 0 then
		quit(0)
	else
		quit(1)
	end if
end if
' Установка 1C:Enterprise
' запуск: im.vbs /install_office_install
'
' @file       1c.vbs
' @copyright  Copyright (c) 2011 iMega ltd. (http://www.imega.ru, info@imega.ru)
' @version    0.1

dim sNameProg: sNameProg = "1C:Enterprise 8.2 (8.2.15.301)"

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

if iVersion = "" then
	includeClass ("iMega_Shell")
	dim oShell: set oShell = new iMega_Shell
	dim resultShell: resultShell = oShell.cmd("msiexec /i " & Chr(34) & "\\data\install\business\1c\bin82\win-8.2.15.301\1CEnterprise 8.2.msi" & Chr(34) & "/qr")
	if resultShell = 0 then
		quit(0)
	else
		quit(1)
	end if
end if
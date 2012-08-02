' Установка 7zip
' запуск: im.vbs /install_7zip
'
' @file       7zip.vbs
' @copyright  Copyright (c) 2011 iMega ltd. (http://www.imega.ru, info@imega.ru)
' @version    0.1

dim sNameProg: sNameProg = "7-zip 9.20*"

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

if oOs.osArchitecture = 64 then
	dim prefix: prefix = "-x64"
end if

if iVersion = "" then
	includeClass ("iMega_Shell")
	dim oShell: set oShell = new iMega_Shell
	dim resultShell: resultShell = oShell.cmd("msiexec /i " & Chr(34) & "\\data\install\archivators\7z920" & prefix & ".msi" & Chr(34) & "/qr")
	if resultShell = 0 then
		quit(0)
	else
		quit(1)
	end if
end if
' Установка/обновление Adobe Reader X (10.1)
' отправить логи на почту
' запуск: im.vbs /install_adobe_reader
'
' @file       reader.vbs
' @copyright  Copyright (c) 2011 iMega ltd. (http://www.imega.ru, info@imega.ru)
' @version    0.3

dim sNameProg: sNameProg = "Adobe Reader*"
dim iCurrentVersion: iCurrentVersion = 167837696

' Если Workstation то продолжаем
includeClass ("iMega_OS"): dim oOs: set oOs = new iMega_OS
if oOs.productType > 1 then quit(0)

includeClass ("iMega_Program"): dim oProg: set oProg = new iMega_Program

dim iVersion: iVersion = 0

with oProg
	.search sNameProg
	iVersion = .version
end with

if iVersion < iCurrentVersion then
	includeClass ("iMega_Shell"): dim oShell: set oShell = new iMega_Shell
	dim resultShell: resultShell = oShell.cmd("\\data.alliance.local\install\viewer\AdbeRdr1010_ru_RU.exe /sPB")
	if resultShell = 0 then
		quit(0)
	else
		quit(1)
	end if
end if
' Установка Java Platform SE binary
' отправить логи на почту
' запуск: im.vbs /install_jre
'
' @file       jre.vbs
' @copyright  Copyright (c) 2011 iMega ltd. (http://www.imega.ru, info@imega.ru)
' @version    0.1

dim sJavaDN: sJavaDN = "Java(TM) 6 Update*"
dim iJavaVer: iJavaVer = 102663596
' Если Workstation то продолжаем
includeClass ("iMega_OS")
dim oOs: set oOs = new iMega_OS
if oOs.productType > 1 then quit(0)

includeClass ("iMega_Program")
dim oProg: set oProg = new iMega_Program

includeClass ("iMega_Shell")
dim oShell: set oShell = new iMega_Shell
dim iVersion: iVersion = 0
with oProg
	.search sJavaDN
	iVersion = .version
end with
echo iVersion & " - " & iJavaVer
if iVersion < iJavaVer or iVersion = null then
	echo "INSTALL"
	dim resultShell: resultShell = oShell.cmd("\\data\install\os\java\jre-6u31-windows-i586-s.exe /s AgreeToLicense=YES IEXPLORER=1 MOZILLA=1 REBOOT=SUPRESS")
	if resultShell = 0 then
		quit(0)
	else
		quit(1)
	end if
end if

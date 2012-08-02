' Adobe Flash Player 10 Plugin
' запуск: im.vbs /install_adobe_flash_player
'
' @file       player.vbs
' @copyright  Copyright (c) 2011 iMega ltd. (http://www.imega.ru, info@imega.ru)
' @version    0.1

dim sCurrentVersion: sCurrentVersion = "11.3.300.268"
dim sNameProg: sNameProg = "Adobe Flash Player 11"
includeClass ("iMega_OS"): dim oOs: set oOs = new iMega_OS
if oOs.productType > 1 then quit(0)

includeClass ("iMega_Program"): dim oProg: set oProg = new iMega_Program
includeClass ("iMega_Shell"): dim oShell: set oShell = new iMega_Shell

dim version
with oProg
	.search sNameProg
	version = .displayVersion & ""
end with

if version <> sCurrentVersion or version = "" then
	dim resultShell: resultShell = oShell.cmd("\\data\install\internet\flash_player\install_flash_player_" & oOs.osArchitecture & "bit.exe -install")
end if
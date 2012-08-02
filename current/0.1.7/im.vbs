' iMega Framework
'
' iMega Framework is based on simplicity, object-oriented best practices,
' facilitates the development and integration of different components
' of a large software project.
'
' Copyright © 2011 Dmitry Gavriloff http://www.imega.ru
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program. If not, see http://www.gnu.org/licenses/.
'
' @file       im.vbs
' @package    iMega_Registry
' @copyright  Copyright (c) 2011 iMega ltd. (http://www.imega.ru, info@imega.ru)
' @license    http://www.imega.ru/license/f4wsh
' @version    0.1.6

Option Explicit

dim classList: set classList = CreateObject("Scripting.Dictionary")

' Include class
' @param nameClass string
' @return void
sub includeClass (nameClass)
	dim prefix
	dim path: path = replace(wscript.scriptFullName, wscript.scriptName, "")
	if instr(nameClass, " ") > 0 then
		prefix = "app"
		nameClass = trim(nameClass)
	else
		prefix = "lib"
	end if
	
	if not classList.exists(nameClass) then
		dim fileToClass: fileToClass = path & prefix & "\" & replace(nameClass, "_", "\") & ".vbs"
		with createObject ("Scripting.FileSystemObject")
			with .openTextFile (fileToClass)
				dim fileData: fileData = .readAll()
				.close
			end with
		end with
		classList.add nameClass, 1
		executeGlobal fileData
	end if
end sub

includeClass ("iMega_Core")
dim item
dim index: index = 0
for each item in split(wscript.arguments(0), "/")
	if index = 1 then
		includeClass item & " "
	end if
	index = index + 1
next


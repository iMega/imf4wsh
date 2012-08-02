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
' @file       fs.vbs
' @package    iMega_FS
' @copyright  Copyright (c) 2011 iMega ltd. (http://www.imega.ru, info@imega.ru)
' @license    http://www.imega.ru/license/f4wsh
' @version    0.1

const READING = 1, _
	WRITING = 2, _
	APPENDING = 8
class iMega_FS
	private oFS, _
		iMode
	private sub Class_Initialize()
		set oFS = CreateObject("Scripting.FileSystemObject")
		iMode = READING
	end sub
	
	private sub Class_Terminate()
        set oFS = nothing
    end sub
	
	'param value string Required. String expression that identifies the folder to create.
	'return Folder Object http://msdn.microsoft.com/en-us/library/1c87day3%28v=VS.85%29.aspx
	public function createDir(value)
		createDir = oFS.CreateFolder(value)
	end function
	
	public function existsDir(value)
	
	end function
	
	public function existsFile(value)
		existsFile = oFS.FileExists(value)
	end function
	
	public function openDir(value)
		if existsDir(value) then
			createDir(value)
		end if
		
	end function
	
	public function openFile(value)
		
	end function

end class
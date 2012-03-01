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
' @file       program.vbs
' @package    iMega_Program
' @copyright  Copyright (c) 2011 iMega ltd. (http://www.imega.ru, info@imega.ru)
' @license    http://www.imega.ru/license/f4wsh
' @version    0.1.6

const MSI = 1, _
	NO_MSI = 0
class iMega_Program
	private oProgram, _
		iTypeInstall, _
		sComments, _
		sDisplayIcon, _
		sDisplayName, _
		sDisplayVersion, _
		iEstimatedSize, _
		sHelpLink, _
		sInstallLocation, _
		iMajorVersion, _
		iMinorVersion, _
		iNoModify, _
		iNoRepair, _
		sProductID, _
		sPublisher, _
		iVersion, _
		sUninstallString, _
		sURLInfoAbout, _
		sURLUpdateInfo
		
	public property get comments() comments = sComments end property
	public property get displayIcon() displayIcon = sDisplayIcon end property
	public property get displayName() displayName = sDisplayName end property
	public property get displayVersion() displayVersion = sDisplayVersion end property
	public property get estimatedSize() estimatedSize = iEstimatedSize end property
	public property get helpLink() helpLink = sHelpLink end property
	public property get installLocation() installLocation = sInstallLocation end property
	public property get majorVersion() majorVersion = iMajorVersion end property
	public property get minorVersion() minorVersion = iMinorVersion end property
	public property get noModify() noModify = iNoModify end property
	public property get noRepair() noRepair = iNoRepair end property
	public property get productID() productID = sProductID end property
	public property get publisher() publisher = sPublisher end property
	public property get typeInstall() typeInstall = iTypeInstall end property
	public property get version() version = iVersion end property
	public property get uninstallString() uninstallString = sUninstallString end property
	public property get urlInfoAbout() urlInfoAbout = sURLInfoAbout end property
	public property get urlUpdateInfo() urlUpdateInfo = sURLUpdateInfo end property

	private property let comments(value) sComments = value end property
	private property let displayIcon(value) sDisplayIcon = value end property
	private property let displayName(value) sDisplayName = value end property
	private property let displayVersion(value) sDisplayVersion = value end property
	private property let estimatedSize(value) iEstimatedSize = value end property
	private property let helpLink(value) sHelpLink = value end property
	private property let installLocation(value) sInstallLocation = value end property
	private property let majorVersion(value) iMajorVersion = value end property
	private property let minorVersion(value) iMinorVersion = value end property
	private property let noModify(value) iNoModify = value end property
	private property let noRepair(value) iNoRepair = value end property
	private property let productID(value) sProductID = value end property
	private property let publisher(value) sPublisher = value end property
	public property let typeInstall(value) iTypeInstall = value	end property
	private property let uninstallString(value) sUninstallString = value end property
	private property let urlInfoAbout(value) sURLInfoAbout = value end property
	private property let urlUpdateInfo(value) sURLUpdateInfo = value end property
	
	private sub Class_Initialize()
		iTypeInstall = NO_MSI
	end sub
	
	private sub Class_Terminate()
        'set objOS = nothing
    end sub
	
	public function search(value)
		includeClass "iMega_Registry"
		dim item, keys, result
		dim registry: set registry = new iMega_Registry
		with registry
			.rootKey = HKEY_LOCAL_MACHINE
			.key = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
			keys = .getKeys
		end with
		
		dim bLike: bLike = false
		if InStr(value, "*") = len(value) then
			value = left(value, len(value) - 1)
			bLike = true
		end if
		
		for each item in keys
			dim found: found = false
			registry.key = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & item
			result = registry.read("DisplayName")
			if bLike then
				if inStr(result, value) = 1 then
					found = true
				end if
			else
				if result = value then
					found = true
				end if
			end if
			
			if found then
				sDisplayName = registry.read("DisplayName")
				sComments = registry.read("Comments")
				sDisplayIcon = registry.read("DisplayIcon")
				sDisplayName = registry.read("DisplayName")
				sDisplayVersion = registry.read("DisplayVersion")
				iEstimatedSize = registry.read("EstimatedSize")
				sHelpLink = registry.read("HelpLink")
				sInstallLocation = registry.read("InstallLocation")
				iMajorVersion = registry.read("MajorVersion")
				iMinorVersion = registry.read("MinorVersion")
				iNoModify = registry.read("NoModify")
				iNoRepair = registry.read("NoRepair")
				sProductID = registry.read("ProductID")
				sPublisher = registry.read("Publisher")
				iVersion = registry.read("Version")
				sUninstallString = registry.read("UninstallString")
				sURLInfoAbout = registry.read("URLInfoAbout")
				sURLUpdateInfo = registry.read("URLUpdateInfo")
			end if
		next
	end function
end class
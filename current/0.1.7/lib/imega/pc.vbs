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
' @file       pc.vbs
' @package    iMega_PC
' @copyright  Copyright (c) 2011 iMega ltd. (http://www.imega.ru, info@imega.ru)
' @license    http://www.imega.ru/license/f4wsh
' @version    0.1

class iMega_PC
	private sName
	
	public function domainRole
		if sName = "" then name
		set oWMISvc = GetObject("winmgmts:" _
			& "{impersonationLevel=impersonate}!\\" _
			& sName & "\root\cimv2")
		set items = oWMISvc.ExecQuery ("Select * from Win32_ComputerSystem")
		for each item in items
			domainRole = item.domainRole
		next
	end function
	
	public function domainRoleMeaning(value)
		select case value
			case 0 domainRoleMeaning = "Standalone Workstation"
			case 1 domainRoleMeaning = "Member Workstation"
			case 2 domainRoleMeaning = "Standalone Server"
			case 3 domainRoleMeaning = "Member Server"
			case 4 domainRoleMeaning = "Backup Domain Controller"
			case 5 domainRoleMeaning = "Primary Domain Controller"
		end select
	end function
	
	public function name
		includeClass ("iMega_OS"): dim os: set os = new iMega_OS
		osType = os.osType
		
		if osType = 16 then
			sName = nameWithAdsi
			if sName = "" then
				sName = nameWithWmi
			end if
		end if
		
		if osType = 17 then
			sName = nameWithWshnet
			if sName = "" then
				sName = nameWithAdsi
				if sName = "" then
					sName = nameWithWmi
				end if
			end if
		end if
		
		if osType = 18 then
			sName = nameWithShell
			if sName = "" then
				sName = nameWithWshnet
				if sName = "" then
					sName = nameWithAdsi
					if sName = "" then
						sName = nameWithWmi
					end if
				end if
			end if
		end if
		
		name = sName
	end function
	
	public function nameDC
		dR = domainRole
		if dR <> 0 or dR <> 2 then
			nameDC = nameWithAD
		end if
	end function
	
	public function nameWithAD
		Set objSysInfo = CreateObject( "ADSystemInfo" )
		nameWithAD = objSysInfo.ComputerName
	end function
	
	public function nameWithShell
		includeClass ("iMega_Shell")
		dim shell: set shell = new iMega_Shell
		nameWithShell = shell.environment("%COMPUTERNAME%")
	end function
	
	public function nameWithWshnet
		set wshNetwork = WScript.CreateObject("WScript.Network")
		nameWithWshnet = wshNetwork.ComputerName
	end function
	
	public function nameWithAdsi
		set objSysInfo = CreateObject("WinNTSystemInfo")
		nameWithAdsi = objSysInfo.ComputerName
	end function
	
	public function nameWithWmi
		Set objWMISvc = GetObject( "winmgmts:\\.\root\cimv2" )
		Set colItems = objWMISvc.ExecQuery("Select * from Win32_ComputerSystem", , 48)
		For Each objItem in colItems
			nameWithWmi = objItem.Name
		Next
	end function
end class
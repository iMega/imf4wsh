' iMega Framework is based on simplicity, object-oriented best practices,
' facilitates the development and integration of different components
' of a large software project.
'
' Copyright (c) 2011 Dmitry Gavriloff http://www.imega.ru
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
' @file       os.vbs
' @package    iMega_OS
' @copyright  Copyright (c) 2011 iMega ltd. (http://www.imega.ru, info@imega.ru)
' @license    http://www.imega.ru/license/f4wsh
' @version    0.2
class iMega_OS
	private objOS, _
		sCaption, _
		sCodeSet, _
		sCountryCode, _
		sLocale, _
		iOsType, _
		iOsArchitecture, _
		pProductType, _
		sServicePack, _
		iSku, _
		iTimeZone, _
		dtUpTime
	
	private sub Class_Initialize()
        pProductType = 0
		set objOS = GetObject("winmgmts:").InstancesOf("Win32_OperatingSystem")
		for each item in objOS
			sCaption = item.caption
			sCodeSet = item.codeSet
			sCountryCode = item.countryCode
			sLocale = item.locale
			iOsArchitecture = getOsArchitecture (item.osArchitecture)
			iOsType = item.osType
			pProductType = item.productType
			sServicePack = item.csdVersion
			'Windows Server 2003, Windows XP, Windows 2000, and Windows NT 4.0:  This property is not available.
			if iOsType > 1 then iSku = item.operatingSystemSKU
			iTimeZone = item.currentTimeZone
			dtUpTime = item.lastBootUpTime
		next
    end sub
	
	private sub Class_Terminate()
        set objOS = nothing
    end sub
    
	public property get caption() caption = sCaption end property
	public property get codeSet() codeSet = sCodeSet end property
	public property get countryCode() countryCode = sCountryCode end property
	public property get locale() locale = sLocale end property
	public property get osArchitecture() osArchitecture = iOsArchitecture end property
	public property get osType() osType = iOsType end property
	public property get productType() productType = pProductType end property
	public property get servicePack() servicePack = sServicePack end property
	public property get sku() sku = iSku end property
	public property get timeZone() timeZone = iTimeZone end property
	'public property get type() type = iOsType end property
	public property get upTime() upTime = dtUpTime end property
	
	public function skuMeaning (value)
		select case value
			case 0 osSKUMeaning = "Undefined"
			case 1 osSKUMeaning = "Ultimate Edition"
			case 2 osSKUMeaning = "Home Basic Edition"
			case 3 osSKUMeaning = "Home Premium Edition"
			case 4 osSKUMeaning = "Enterprise Edition"
			case 5 osSKUMeaning = "Home Basic N Edition"
			case 6 osSKUMeaning = "Business Edition"
			case 7 osSKUMeaning = "Standard Server Edition"
			case 8 osSKUMeaning = "Datacenter Server Edition"
			case 9 osSKUMeaning = "Small Business Server Edition"
			case 10 osSKUMeaning = "Enterprise Server Edition"
			case 11 osSKUMeaning = "Starter Edition"
			case 12 osSKUMeaning = "Datacenter Server Core Edition"
			case 13 osSKUMeaning = "Standard Server Core Edition"
			case 14 osSKUMeaning = "Enterprise Server Core Edition"
			case 15 osSKUMeaning = "Enterprise Server Edition for Itanium-Based Systems"
			case 16 osSKUMeaning = "Business N Edition"
			case 17 osSKUMeaning = "Web Server Edition"
			case 18 osSKUMeaning = "Cluster Server Edition"
			case 19 osSKUMeaning = "Home Server Edition"
			case 20 osSKUMeaning = "Storage Express Server Edition"
			case 21 osSKUMeaning = "Storage Standard Server Edition"
			case 22 osSKUMeaning = "Storage Workgroup Server Edition"
			case 23 osSKUMeaning = "Storage Enterprise Server Edition"
			case 24 osSKUMeaning = "Server For Small Business Edition"
			case 25 osSKUMeaning = "Small Business Server Premium Edition"
		end select
	end function
	
	public function osTypeMeaning (value)
		select case value
			case 0 osTypeMeaning = "Unknown"
			case 1 osTypeMeaning = "Other"
			case 2 osTypeMeaning = "MACROS"
			case 3 osTypeMeaning = "ATTUNIX"
			case 4 osTypeMeaning = "DGUX"
			case 5 osTypeMeaning = "DECNT"
			case 6 osTypeMeaning = "Digital UNIX"
			case 7 osTypeMeaning = "OpenVMS"
			case 8 osTypeMeaning = "HPUX"
			case 9 osTypeMeaning = "AIX"
			case 10 osTypeMeaning = "MVS"
			case 11 osTypeMeaning = "OS400"
			case 12 osTypeMeaning = "OS/2"
			case 13 osTypeMeaning = "JavaVM"
			case 14 osTypeMeaning = "MSDOS"
			case 15 osTypeMeaning = "WIN3x"
			case 16 osTypeMeaning = "WIN95"
			case 17 osTypeMeaning = "WIN98"
			case 18 osTypeMeaning = "WINNT"
			case 19 osTypeMeaning = "WINCE"
			case 20 osTypeMeaning = "NCR3000"
			case 21 osTypeMeaning = "NetWare"
			case 22 osTypeMeaning = "OSF"
			case 23 osTypeMeaning = "DC/OS"
			case 24 osTypeMeaning = "Reliant UNIX"
			case 25 osTypeMeaning = "SCO UnixWare"
			case 26 osTypeMeaning = "SCO OpenServer"
			case 27 osTypeMeaning = "Sequent"
			case 28 osTypeMeaning = "IRIX"
			case 29 osTypeMeaning = "Solaris"
			case 30 osTypeMeaning = "SunOS"
			case 31 osTypeMeaning = "U6000"
			case 32 osTypeMeaning = "ASERIES"
			case 33 osTypeMeaning = "TandemNSK"
			case 34 osTypeMeaning = "TandemNT"
			case 35 osTypeMeaning = "BS2000"
			case 36 osTypeMeaning = "LINUX"
			case 37 osTypeMeaning = "Lynx"
			case 38 osTypeMeaning = "XENIX"
			case 39 osTypeMeaning = "VM/ESA"
			case 40 osTypeMeaning = "Interactive UNIX"
			case 41 osTypeMeaning = "BSDUNIX"
			case 42 osTypeMeaning = "FreeBSD"
			case 43 osTypeMeaning = "NetBSD"
			case 44 osTypeMeaning = "GNU Hurd"
			case 45 osTypeMeaning = "OS9"
			case 46 osTypeMeaning = "MACH Kernel"
			case 47 osTypeMeaning = "Inferno"
			case 48 osTypeMeaning = "QNX"
			case 49 osTypeMeaning = "EPOC"
			case 50 osTypeMeaning = "IxWorks"
			case 51 osTypeMeaning = "VxWorks"
			case 52 osTypeMeaning = "MiNT"
			case 53 osTypeMeaning = "BeOS"
			case 54 osTypeMeaning = "HP MPE"
			case 55 osTypeMeaning = "NextStep"
			case 56 osTypeMeaning = "PalmPilot"
			case 57 osTypeMeaning = "Rhapsody"
		end select
	end function
	
	private function getOsArchitecture (value)
		if (item.OSArchitecture = "32-bit") then
			getOsArchitecture = 32
		else 
			getOsArchitecture = 64
		end if
	end function
end class
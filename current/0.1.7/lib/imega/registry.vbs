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
' @file       registry.vbs
' @package    iMega_Registry
' @copyright  Copyright (c) 2011 iMega ltd. (http://www.imega.ru, info@imega.ru)
' @license    http://www.imega.ru/license/f4wsh
' @version    0.1.6

const HKEY_CLASSES_ROOT = &H80000000, _
	HKEY_CURRENT_USER = &H80000001, _
	HKEY_LOCAL_MACHINE = &H80000002, _
	HKEY_USERS = &H80000003, _
	HKEY_CURRENT_CONFIG = &H80000005, _
	HKEY_DYN_DATA = &H80000006
	
const KEY_QUERY_VALUE = &H0001, _
	KEY_SET_VALUE = &H0002, _
	KEY_CREATE_SUB_KEY = &H0004, _
	DELETE = &H00010000

const REG_SZ = 1, _
	REG_EXPAND_SZ = 2, _
	REG_BINARY = 3, _
	REG_DWORD = 4, _
	REG_QWORD = 11, _
	REG_MULTI_SZ = 7

const SUCCESS = 0, _
	KEY_NOT_EXIST = 2, _
	PARAM_NOT_EXIST = 3, _
	ACCESS_DENIED = 5
	
class iMega_Registry
	private oRegistry, _
		oLog, _
		bAccessGet, _
		bAccessSet, _
		bAccessSub, _
		bAccessDel, _
		iCurrentParamType, _
		iRootKey, _
		sKey, _
		aKeys, _
		sPCName
	private sub Class_Initialize()
		iRootKey = HKEY_CURRENT_USER
		sPCName = "."
		set oRegistry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
			sPCName & "\root\default:StdRegProv")
	end sub
	
	private sub Class_Terminate()
        set objOS = nothing
    end sub
	
	public property get rootKey() rootKey = iRootKey end property
	public property get key() key = sKey end property
	public property get pcName() pcName = sPCName end property
	
	public property let key(value)
		value = splashTrim(value)
		sKey = value
		bAccessGet = checkAccess(KEY_QUERY_VALUE)
		bAccessSet = checkAccess(KEY_SET_VALUE)
		bAccessSub = checkAccess(KEY_CREATE_SUB_KEY)
		bAccessDel = checkAccess(DELETE)
	end property
	public property let setLog(value) set oLog = value end property
	public property let pcName(value)
		sPCName = value
		set oRegistry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
			sPCName & "\root\default:StdRegProv")
	end property
	public property let rootKey(value) iRootKey = value end property
	
	private function getBinary(param)
		result = oRegistry.GetBinaryValue (iRootKey, sKey, param, value)
		if result = SUCCESS then
			log INFO, "SUCCESS getBinary " & param
			result = value
		end if
		getBinary = result
	end function
	
	public function getKeys()
		if bAccessGet = true then
			oRegistry.EnumKey iRootKey, sKey, aKeys
			result = aKeys
		else
			log ERROR, "ACCESS_DENIED getKeys " & sKey
			result = ACCESS_DENIED
		end if
		getKeys = result
	end function
	
	public function getParams()
		if bAccessGet = true then
			result = oRegistry.EnumValues (iRootKey, sKey, aParams, aParamsTypes)
			if result = SUCCESS then
				log INFO, "SUCCESS getParams " & sKey
				result = aParams
			end if
		else
			log ERROR, "ACCESS_DENIED getParams " & sKey
			result = ACCESS_DENIED
		end if
		getParams = result
	end function
	
	private function getExString(param)
		result = oRegistry.GetExpandedStringValue (iRootKey, sKey, param, value)
		if result = SUCCESS then
			log INFO, "SUCCESS getExString " & param
			result = value
		end if
		getExString = result
	end function
	
	private function getMuString(param)
		result = oRegistry.GetMultiStringValue (iRootKey, sKey, param, value)
		if result = SUCCESS then
			log INFO, "SUCCESS getMuString " & param
			result = value
		end if
		getMuString = result
	end function
	
	private function getString(param)
		result = oRegistry.GetStringValue (iRootKey, sKey, param, value)
		if result = SUCCESS then
			log INFO, "SUCCESS getString " & param
			result = value
		end if
		getString = result
	end function
	
	private function getDWord(param)
		result = oRegistry.GetDWORDValue (iRootKey, sKey, param, value)
		if result = SUCCESS then
			log INFO, "SUCCESS getDWord " & param
			result = value
		end if
		getDWord = result
	end function
	
	private function getQWord(param)
		result = oRegistry.GetQWORDValue (iRootKey, sKey, param, value)
		if result = SUCCESS then
			log INFO, "SUCCESS getQWord " & param
			result = value
		end if
		getQWord = result
	end function
	
	private function checkAccess(access)
		on error resume next
		result = oRegistry.CheckAccess (iRootKey, sKey, access, bGranted)
		if Err.Number <> 0 then
			log ERROR, "checkAccess: " & " " & Err.Description
			wscript.quit(1)
		end if
		if result = SUCCESS then
			log INFO, "SUCCESS checkAccess " & sKey
			result = bGranted
		else
			log WARN, "WARNING checkAccess: " & sKey & ". Code: " & result
		end if
		checkAccess = result
	end function
	
	public function existKey(value)
		if bAccessGet = true then
			log WARN, "KEY_NOT_EXIST existKey " & value
			result = KEY_NOT_EXIST
			dim aKeys: aKeys = getKeys()
			for index = 0 to UBound(aKeys)
				if aKeys(index) = value then
					result = true
				end if
			next
		else
			log WARN, "ACCESS_DENIED existKey " & value
			result = ACCESS_DENIED
		end if
		existKey = result
	end function
	
	public function existParam(value)
		if bAccessGet = true then
			dim result, aParams, aParamsTypes, param
			result = oRegistry.EnumValues (iRootKey, sKey, aParams, aParamsTypes)
			if result = SUCCESS then
				result = PARAM_NOT_EXIST
				if isArray(aParams) = true then
					for index = 0 to UBound(aParams)
						if aParams(index) = value then
							result = true
							iCurrentParamType = aParamsTypes(index)
						end if
					next
				end if
				if result = PARAM_NOT_EXIST then log WARN, "PARAM_NOT_EXIST existParam " & value
			else
				log WARN, "KEY_NOT_EXIST existParam " & value
				result = KEY_NOT_EXIST
			end if
		else
			log WARN, "ACCESS_DENIED existParam " & value
			result = ACCESS_DENIED
		end if
		existParam = result
	end function
	
	private function log(iType, message)
		if isObject(oLog) then
			oLog.log iType, "class: iMega_Registry" & chr(10)+chr(13) & message
		end if
	end function
	
	'error codes
	' 2 - key not exist
	' 3 - param not exist
	' 5 - access denied
	public function read(param)
		if bAccessGet = true then
			dim result: result = existParam(param)
			if result = true then
				select case iCurrentParamType
					case REG_SZ
						read = getString(param)
					case REG_EXPAND_SZ
						read = getExString(param)
					case REG_BINARY
						read = getBinary(param)
					case REG_DWORD
						read = getDWord(param)
					case REG_QWORD
						read = getQWord(param)
					case REG_MULTI_SZ
						read = getMuString(param)
				end select
			else
				log WARN, "PARAM_NOT_EXIST OR KEY_NOT_EXIST read " & param
				read = result
			end if
		else
			log WARN, "ACCESS_DENIED read " & param
			read = bAccessGet
		end if
	end function
	
	private function setString(param, value)
		setString = oRegistry.SetStringValue(iRootKey, sKey, param, value)
	end function
	
	private function setExString(param, value)
		setExString = oRegistry.SetExpandedStringValue(iRootKey, sKey, param, value)
	end function
	
	private function setBinary(param, value)
		setBinary = oRegistry.SetBinaryValue(iRootKey, sKey, param, value)
	end function
	
	private function setDWord(param, value)
		setDWord = oRegistry.SetDWORDValue(iRootKey, sKey, param, value)
	end function
	
	private function setQWord(param, value)
		setQWord = oRegistry.SetQWORDValue(iRootKey, sKey, param, value)
	end function
	
	private function setMuString(param, value)
		setMuString = oRegistry.SetMultiStringValue(iRootKey, sKey, param, value)
	end function
	
	private function splashTrim(value)
		if left(value, 1) = "\" then
			value = mid(value, 2, len(value)-1)
		end if
		splashTrim = value
	end function
	
	'error codes
	' 5 - access denied
	public function write(param, typeParam, value)
		if bAccessSet = true then
			param = splashTrim(param)
			dim newPath, realParam
			isPath = instr(param, "\")
			if isPath > 0 then
				dim splitValue: splitValue = split(param, "\")
				realParam = splitValue(ubound(splitValue))
				newPath = mid(param, 1, len(param) - len(realParam) - 1)
				if bAccessSub = true then
					result = oRegistry.CreateKey(iRootKey, sKey & "\" & newPath)
					if result = SUCCESS then
						log INFO, "CreateKey " & sKey & "\" & newPath
						sKey = sKey & "\" & newPath
					else
						log ERROR, "CreateKey " & sKey & "\" & newPath
						write = result
					end if
				else
					log WARN, "ACCESS_DENIED " & sKey
					write = bAccessSub
				end if
			else
				realParam = param
			end if
			select case typeParam
				case REG_SZ
					write = setString(realParam, value)
				case REG_EXPAND_SZ
					write = setExString(realParam, value)
				case REG_BINARY
					write = setBinary(realParam, value)
				case REG_DWORD
					write = setDWord(realParam, value)
				case REG_QWORD
					write = setQWord(realParam, value)
				case REG_MULTI_SZ
					write = setMuString(realParam, value)
			end select
		else
			log WARN, "ACCESS_DENIED write " & param
			write = bAccessSet
		end if
	end function
end class
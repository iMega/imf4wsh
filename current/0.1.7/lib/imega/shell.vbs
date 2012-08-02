' iMega F4WSH
'
' LICENSE
'
' not many words about the license :)
'
' @file       shell.vbs
' @category   F4WSH
' @package    iMega_Shell
' @copyright  Copyright (c) 2011 iMega ltd. (http://www.imega.ru, info@imega.ru)
' @license    http://www.imega.ru/license/f4wsh
' @version    0.1

class iMega_Shell
	private oShell, _
		iWindowStyle
	
	'Constructor
	private sub Class_Initialize()
		set oShell = WScript.CreateObject("WScript.Shell")
	end sub
	
	'Destructor
	private sub Class_Terminate()
        set oShell = nothing
    end sub
	
	'getObject		Returns an instance of an object in the class initialization
	public property get getObject() set getObject = oShell end property
	
	'windowStyle	Integer value indicating the appearance of the program's window.
	'				Note that not all programs make use of this information.
	public property get windowStyle() windowStyle = iWindowStyle end property
	public property let windowStyle(value) iWindowStyle = value end property
	
	'Runs a program in a new process
	'
	'@param		string	value	String value indicating the command line you 
	'							want to run. You must include any parameters
	'							you want to pass to the executable file.
	'@return	int		
	public function cmd (value)
		dim sCmd : sCmd = "cmd /c start /wait " & value
		cmd = oShell.Run (sCmd, 0, true)
	end function
	
	public function environment (value)
		environment = oShell.ExpandEnvironmentStrings(value)
	end function
end class
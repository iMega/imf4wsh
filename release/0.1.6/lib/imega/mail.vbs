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
' @file       mail.vbs
' @package    iMega_Mail
' @copyright  Copyright (c) 2011 iMega ltd. (http://www.imega.ru, info@imega.ru)
' @license    http://www.imega.ru/license/f4wsh
' @version    0.1.6
 
'cdoSendUsingPickup  Send message using the local SMTP service pickup directory.
'cdoSendUsingPort    Send the message using the network (SMTP over the network).
const cdoSendUsingPickup = 1, _
	cdoSendUsingPort = 2

'Authentication
'AUTH_ANONYMOUS  Do not authenticate
'AUTH_BASIC      basic (clear-text) authentication
'AUTH_NTLM       NTLM authentication protocol
const cdoAnonymous = 0, _
	cdoBasic = 1, _
	cdoNTLM = 2

'Delivery Status Notifications
'cdoDSNDefault             No delivery status notifications are issued.
'cdoDSNNever               No delivery status notifications are issued.
'cdoDSNFailure             Returns a delivery status notification if delivery fails.
'cdoDSNSuccess             Returns a delivery status notification if delivery succeeds.
'cdoDSNDelay               Returns a delivery status notification if delivery is delayed.
'cdoDSNSuccessFailOrDelay  Returns a delivery status notification if delivery succeeds, fails, or is delayed.
const cdoDSNDefault = 0, _
	cdoDSNNever = 1, _
	cdoDSNFailure = 2, _
	cdoDSNSuccess = 4, _
	cdoDSNDelay = 8, _
	cdoDSNSuccessFailOrDelay = 14

'Fields collection
const cdoSendUsingMethod = "http://schemas.microsoft.com/cdo/configuration/sendusing", _
	cdoSMTPServer = "http://schemas.microsoft.com/cdo/configuration/smtpserver", _
	cdoSMTPServerPort = "http://schemas.microsoft.com/cdo/configuration/smtpserverport", _
	cdoSMTPConnectionTimeout = "http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout", _
	cdoSMTPAuthenticate = "http://schemas.microsoft.com/cdo/configuration/smtpauthenticate", _
	cdoSendUserName = "http://schemas.microsoft.com/cdo/configuration/sendusername", _
	cdoSendPassword = "http://schemas.microsoft.com/cdo/configuration/sendpassword"

class iMega_Mail
	private oMail, _
		oConf, _
		sAttachment, _
		iAuth, _
		sFrom, _
		sMessage, _
		sPass, _
		iPort, _
		sRecipient, _
		iSendMethod, _
		sServer, _
		sSubject, _
		iTimeout, _
		sUser
	
	public property get attachment() attachment = sAttachment end property
	public property get auth() auth = iAuth end property
	public property get from() from = sFrom end property
	public property get message() message = sMessage end property	
	public property get pass() pass = sPass end property
	public property get port() port = iPort end property
	public property get recipient() recipient = sRecipient end property
	public property get sendMethod() sendMethod = iSendMethod end property
	public property get server() server = sServer end property
	public property get subject() subject = sSubject end property
	public property get timeout() timeout = iTimeout end property
	public property get user() user = sUser end property
	
	public property let attachment(value) sAttachment = value end property
	public property let auth(value) iAuth = value end property
	public property let from(value) sFrom = value end property
	public property let message(value) sMessage = value end property	
	public property let pass(value) sPass = value end property
	public property let port(value) iPort = value end property
	public property let recipient(value) sRecipient = value end property
	public property let sendMethod(value) iSendMethod = value end property
	public property let server(value) sServer = value end property
	public property let subject(value) sSubject = value end property
	public property let timeout(value) iTimeout = value end property
	public property let user(value) sUser = value end property
	
	'Constructor
	private sub Class_Initialize()
		set oMail = CreateObject("CDO.Message")
		iPort = 25
		iSendMethod = cdoSendUsingPort
		'set oConf = CreateObject("CDO.Message")
	end sub
	
	'Destructor
	private sub Class_Terminate()
        set oMail = nothing
		'set oConf = nothing
    end sub
	
	'Send message
	public function send()
		with oMail
			.Subject = sSubject
			.From = sFrom
			.To = sRecipient
			.TextBody = sMessage
			with .Configuration.Fields
				.Item (cdoSendUsingMethod) = iSendMethod
				.Item (cdoSMTPServer) = sServer
				.Item (cdoSMTPServerPort) = iPort
				.Update
			end with
			.Send
		end with
	end function
end class
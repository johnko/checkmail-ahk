/*
	I have modified code from / Thanks to:
		Lexikos - Minimize to Tray - http://www.autohotkey.com/forum/topic18260.html
		daonlyfreez - Winsock 2 SMTP/POP - http://www.autohotkey.com/forum/topic16184.html
*/

/* Example calls...
SMTP (port 25):
	HELO
	MAIL FROM:<someone@someserver.com>
	RCPT TO:<someone@someotherserver.com>
	DATA
	From: "Some One" <someone@someserver.com>
	To: <someone@someotherserver.com>
	Subject: Hi there, how are you?
	"Message body"
	.
	QUIT
POP (port 110)
	HELO
	USER someone@someserver.com
	PASS password
	STAT
	LIST
	TOP 1 10
	QUIT
IRC (port 6667)
	/list /list -*ahk*
	/join #channel
	hi there
	/quit
*/
	#SingleInstance,Force 
	#NoEnv 
	;#NoTrayIcon 
	#Persistent 
	SetBatchLines,-1
	Remote_Address_Org = ; // nothing
	Remote_Port = ; // nothing
	Remote_U := ; // nothing
	Remote_P := ; // nothing
	Remote_Time = 300 ; // will be multiplied by 60000
	;----------------

	Gosub, TrayMenu
	;----------------

	Gui, Font, bold
	Gui, Add, Text, x10 y10, CheckMail
	Gui, Font, norm
	Gui, Add, Text, x73 y10, - v0.100516 by
	Gui, Font, cBlue underline
	Gui, Add, Text, x150 y10 gLaunchWebsite, johnko.ca
	Gui, Font, cBlack norm
	Gui, Add, Text, x10 y36, Server Type*
	Gui, Add, DropDownList, x10 y52 w70 vGType, IMAP|POP3
	Gui, Add, Text, x90 y36, Host*
	Gui, Add, Edit, x90 y52 w110 R1 vGServer, %Remote_Address_Org%
	Gui, Add, Text, x210 y36, Port
	Gui, Add, Edit, x210 y52 w50 R1 vGPort, %Remote_Port%
	Gui, Add, Text, x10 y81, Username:
	Gui, Add, Edit, x70 y78 w190 R1 vGUser, %Remote_U%
	Gui, Add, Text, x10 y107, Password:
	Gui, Add, Edit, x70 y104 w190 R1 vGPass, %Remote_P%
	Gui, Add, Text, x10 y133, Check every
	Gui, Add, DropDownList, x75 y130 w40 vGTime, 5||1|10|20|30
	Gui, Add, Text, x119 y133, minute(s).
	Gui, Add, Button, x10 y156 w60 gTestSettings, Test
	Gui, Add, Progress, x80 y159 w110 h18 -Smooth +0x8 vMyProgress
	Gui, Add, Button, x200 y156 w60 gSaveSettings, Save
	Gui, Add, Text, x10 y186 w250, NOTE: This is public domain, open source software. Proceed with caution and use at your own risk!
	; // Gui, Add, Button, x130 y160 w60 h24 gExitSub, Exit
	Gui, Show, , CheckMail - Settings
	LinesReceived:=0
	OnExit, ExitSub  ; // For connection cleanup purposes.
return
;----------------

TrayMenu:
Menu, Tray, Tip, CheckMail by johnko.ca
; Menu, Tray, MainWindow
Menu, Tray, NoStandard
Menu, Tray, DeleteAll
Menu, Tray, Add, Settings, GuiShowHide
Menu, Tray, Default, Settings
Menu, Tray, Click, 1
Menu, Tray, Add, Exit, ExitSub
return
;----------------

LaunchWebsite:
Run, http://www.johnko.ca/checkmail
return
;----------------

SaveSettings:
; // TODO if test passed, enable save
; // SetTimer, CheckIMAPMail, % Remote_Time * 60000
return
;----------------

BarPush:
GuiControl, , MyProgress, +1
return
;----------------

TestSettings:
GuiControl, Disable, Test
SetTimer, BarPush, 45
Gosub, ValidateInput
if (goodsettings = 1)
{
	Gosub, CheckIMAPMail
	SetTimer, EnableTestButton, -3000
}
else
	SetTimer, EnableTestButton, -1
return
;----------------

EnableTestButton:
SetTimer, BarPush, Off
GuiControl, Enable, Test
GoSub, ResetVars
return
;----------------

ResetVars:
socket := ; //nothing
remotesocket := ; //nothing
return
;----------------

ValidateInput:
goodsettings = 0
GuiControlGet, Remote_Type, , GType
GuiControlGet, Remote_Address_Org, , GServer
GuiControlGet, Remote_Port, , GPort
GuiControlGet, Remote_U, , GUser
GuiControlGet, Remote_P, , GPass
GuiControlGet, Remote_Time, , GTime
; // if port empty, but type selected, autofill port
if (Remote_Type <> "")
{
	if (Remote_Address_Org <> "")
	{
		if (Remote_Port = "")
		{
			if (Remote_Type = "IMAP")
			{
				Remote_Port = 143
				GuiControl,, GPort, 143
				goodsettings = 1
			}
			else if (Remote_Type = "POP3")
			{
				Remote_Port = 110
				GuiControl,, GPort, 110
				goodsettings = 1
			}
		}
		else
			goodsettings = 1
	}
	else
		MsgBox, 48, Error, Missing required setting "Host".
}
else
	MsgBox, 48, Error, Missing required setting "Server Type".
return
;----------------

CheckIMAPMail:
SetTimer, Connection_Init, -500
SetTimer, Login, -1000
SetTimer, CheckUnseen, -1500
SetTimer, Logout, -2000
SetTimer, CleanUp, -2500
return
;----------------

Login:
SendData(remotesocket,". LOGIN " . Remote_U . " " . Remote_P . "`r`n")
; // LinesReceived++
; // ShowReceived = %ShowReceived%`n%LinesReceived%: . LOGIN %Remote_U% %Remote_P%
; // GuiControl,, MyEdit, %ShowReceived%
return
;----------------

CheckUnseen:
SendData(remotesocket,". STATUS INBOX (UNSEEN)" . "`r`n")
; // LinesReceived++
; // ShowReceived = %ShowReceived%`n%LinesReceived%: . STATUS INBOX (UNSEEN)
; // GuiControl,, MyEdit, %ShowReceived%
return
;----------------

Logout:
SendData(remotesocket,". LOGOUT" . "`r`n")
; // LinesReceived++
; // ShowReceived = %ShowReceived%`n%LinesReceived%: . LOGOUT
; // GuiControl,, MyEdit, %ShowReceived%
return
;----------------

CleanUp:
DllCall("Ws2_32\WSACleanup") 
return
;----------------

Connection_Init:
	StringReplace, Remote_Address_Stripped, Remote_Address_Org, `., , A ; // Get IP if domain
	If Remote_Address_Stripped is not number
	{
		IPs := HostToIp(Remote_Address_Org)
		DllCall("Ws2_32\WSACleanup") ; // always inlude this line after calling to release the socket connection
		if IPs <> -1 ; no error occurred
		{
			If IPs contains `n
			{
				StringSplit, anIP, IPs, `n
				; // using first ip found
				Remote_Address := anIP1
			}
			Else
				Remote_Address := IPs
		}
		else
		{
			MsgBox, 48, Error, Host "%Remote_Address_Org%" not found`, Exiting...
			Gosub, ExitSub
		}
	}
	Else
		Remote_Address := Remote_Address_Org
	remotesocket := ConnectToAddress(Remote_Address, Remote_Port)
	if remotesocket = -1  ; // Connection failed (it already displayed the reason).
		ExitApp
	; // Find this script's main window:
	Process, Exist  ; // This sets ErrorLevel to this script's PID (it's done this way to support compiled scripts).
	DetectHiddenWindows On
	ScriptMainWindowId := WinExist("ahk_class AutoHotkey ahk_pid " . ErrorLevel)
	DetectHiddenWindows Off
	; // When the OS notifies the script that there is incoming data waiting to be received,
	; // the following causes a function to be launched to read the data:
	NotificationMsg = 0x5555  ; // An arbitrary message number, but should be greater than 0x1000.
	OnMessage(NotificationMsg, "ReceiveData")
	; // Set up the connection to notify this script via message whenever new data has arrived.
	; // This avoids the need to poll the connection and thus cuts down on resource usage.
	FD_READ = 1     ; // Received when data is available to be read.
	FD_CLOSE = 32   ; // Received when connection has been closed.
	if DllCall("Ws2_32\WSAAsyncSelect", "UInt", remotesocket, "UInt", ScriptMainWindowId, "UInt", NotificationMsg, "Int", FD_READ|FD_CLOSE)
	{
		MsgBox, 48, Error, % "WSAAsyncSelect() indicated Winsock error " . DllCall("Ws2_32\WSAGetLastError")
		ExitApp
	}
	/*
	; // The following Loop was causing problems
	Loop ; // Wait for incomming connections
	{
		; // accept requests that are in the pipeline of the socket
		conncheck := DllCall("Ws2_32\accept", "UInt", remotesocket, "UInt", &SocketAddress, "Int", SizeOfSocketAddress)
		; // Ws2_22/accept returns the new Connection-Socket if a connection request was in the pipeline
		; // on failure it returns an negative value
		if conncheck > 1
		{
			MsgBox Incoming connection accepted
			break
		}
		sleep 500 ; // wait half 1 second then accept again
	}
	*/
return
;----------------

ConnectToAddress(IPAddress, Port)
{
	global socket
	; // Returns -1 (INVALID_SOCKET) upon failure or the socket ID upon success.
    VarSetCapacity(wsaData, 32)  ; // The struct is only about 14 in size, so 32 is conservative.
    result := DllCall("Ws2_32\WSAStartup", "UShort", 0x0002, "UInt", &wsaData) ; // Request Winsock 2.0 (0x0002)
    ; // Since WSAStartup() will likely be the first Winsock function called by this script,
    ; // check ErrorLevel to see if the OS has Winsock 2.0 available:
    if ErrorLevel
    {
        MsgBox, 48, Error, WSAStartup() could not be called due to error %ErrorLevel%. Winsock 2.0 or higher is required.
        return -1
    }
    if result  ; // Non-zero, which means it failed (most Winsock functions return 0 upon success).
    {
        MsgBox, 48, Error, % "WSAStartup() indicated Winsock error " . DllCall("Ws2_32\WSAGetLastError")
        return -1
    }
    AF_INET = 2
    SOCK_STREAM = 1
    IPPROTO_TCP = 6
    socket := DllCall("Ws2_32\socket", "Int", AF_INET, "Int", SOCK_STREAM, "Int", IPPROTO_TCP)
    if socket = -1
    {
        MsgBox, 48, Error, % "socket() indicated Winsock error " . DllCall("Ws2_32\WSAGetLastError")
        return -1
    }
    ; // Prepare for connection:
    SizeOfSocketAddress = 16
    VarSetCapacity(SocketAddress, SizeOfSocketAddress)
    InsertInteger(2, SocketAddress, 0, AF_INET)   ; sin_family
    InsertInteger(DllCall("Ws2_32\htons", "UShort", Port), SocketAddress, 2, 2)   ; sin_port
    InsertInteger(DllCall("Ws2_32\inet_addr", "Str", IPAddress), SocketAddress, 4, 4)   ; sin_addr.s_addr
    ; // Attempt connection:
    if DllCall("Ws2_32\connect", "UInt", socket, "UInt", &SocketAddress, "Int", SizeOfSocketAddress)
    {
        MsgBox, 48, Error, % "connect() indicated Winsock error " . DllCall("Ws2_32\WSAGetLastError") . "?"
        return -1
    }
    return socket  ; // Indicate success by returning a valid socket ID rather than -1.
}
;----------------

SendData(wParam,SendData, Repeat=1, Delay=0)
{
	socket := wParam
	; //  SendDataSize := VarSetCapacity(SendData)
	; //  SendDataSize += 1
	Loop % Repeat
	{
		SendIt := SendData . "(" . A_Index . ")"
		SendDataSize := VarSetCapacity(SendIt)
		SendDataSize += 1
		sendret := DllCall("Ws2_32\send", "UInt", socket, "Str", SendIt, "Int", strlen(SendIt), "Int", 0)
		WinsockError := DllCall("Ws2_32\WSAGetLastError")
		if WinsockError <> 0 ; // WSAECONNRESET, which happens when Network closes via system shutdown/logoff.
			; // Since it's an unexpected error, report it.  Also exit to avoid infinite loop.
			MsgBox, 48, Error, % "send() indicated Winsock error " . WinsockError
		sleep, Delay
	}
; // send( sockConnected,> welcome, strlen(welcome) + 1, NULL);
}
;----------------

ReceiveData(wParam, lParam)
; // By means of OnMessage(), this function has been set up to be called automatically whenever new data
; // arrives on the connection.
{
	Critical
	global ShowReceived
	global MyEdit
	global LinesReceived
	socket := wParam
	ReceivedDataSize = 4096  ; // Large in case a lot of data gets buffered due to delay in processing previous data.
	VarSetCapacity(ReceivedData, ReceivedDataSize, 0)  ; // 0 for last param terminates string for use with recv().
	ReceivedDataLength := DllCall("Ws2_32\recv", "UInt", socket, "Str", ReceivedData, "Int", ReceivedDataSize, "Int", 0)
	if ReceivedDataLength = 0
	{
		; // The connection was gracefully closed
		TrayTip, New Mail, % ShowReceived ; // FoundMail
		return 1 ; // ExitApp  ; // The OnExit routine will call WSACleanup() for us.
	}
	if ReceivedDataLength = -1
	{
		WinsockError := DllCall("Ws2_32\WSAGetLastError")
		if WinsockError = 10035  ; // WSAEWOULDBLOCK, which means "no more data to be read".
			return 1  ; // Should probably never happen since we were notified there is data on the connection, yet now we're told there's none?
		if WinsockError <> 10054 ; //  WSAECONNRESET, which happens when Network closes via system shutdown/logoff.
			; // Since it's an unexpected error, report it.  Also exit to avoid infinite loop.
			MsgBox, 48, Error, % "recv() indicated Winsock error " . WinsockError
		ExitApp  ; // The OnExit routine will call WSACleanup() for us.
	}
	; // Since above didn't return or exit, process the data that was just received.
	Loop  ; // For each binary-zero-delimited segment in the data.
	{
		Loop, parse, ReceivedData, `n, `r  ; For each line in this segment.
		{
			LinesReceived++
			if (LinesReceived = 1) {
				ShowReceived = %A_LoopField% ; // %LinesReceived%: %A_LoopField%
			} else {
				ShowReceived = %ShowReceived%`n%A_LoopField% ; // %ShowReceived%`n%LinesReceived%: %A_LoopField%
			}
			; // Tooltip % ShowReceived
		}
		ReceivedDataLengthApparent := strlen(ReceivedData)
		if (ReceivedDataLength-1 <= ReceivedDataLengthApparent)  ; // -1 to adjust for the legitimate/last zero-termintor at the end of the last segment.
			break   ; // No more binary-zero-delimited segements are present.
		; // Otherwise, there's a binary zero "hiding" more data that lies to its right.
		DllCall("RtlMoveMemory", str, ReceivedData  ; // Shift the data leftward to eliminate from consideration the segement that was just processed.
			, UInt, &ReceivedData + ReceivedDataLengthApparent + 1
			, UInt, ReceivedDataLength - ReceivedDataLengthApparent)
		ReceivedDataLength -= ReceivedDataLengthApparent + 1  ; // Adjust length to reflect actual NEW length of ReceivedData.
	}
	return 1  ; // Tell the program that no further processing of this message is needed.
}
;----------------

HostToIp(NodeName) ; // returns -1 if unsuccessfull or a newline seperated list of valid IP addresses on success
{
	VarSetCapacity(wsaData, 32)  ; // The struct is only about 14 in size, so 32 is conservative.
	result := DllCall("Ws2_32\WSAStartup", "UShort", 0x0002, "UInt", &wsaData) ; Request Winsock 2.0 (0x0002)
	if ErrorLevel   ; // check ErrorLevel to see if the OS has Winsock 2.0 available:
	{
		MsgBox, 48, Error, WSAStartup() could not be called due to error %ErrorLevel%. Winsock 2.0 or higher is required.
		return -1
	}
	if result  ; // Non-zero, which means it failed (most Winsock functions return 0 on success).
	{
		MsgBox, 48, Error, % "WSAStartup() indicated Winsock error " . DllCall("Ws2_32\WSAGetLastError") ; %
		return -1
	}
	PtrHostent := DllCall("Ws2_32\gethostbyname", str, Nodename)
	if (PtrHostent = 0)
		return -1
	VarSetCapacity(hostent,16,0)
	DllCall("RtlMoveMemory",UInt,&hostent,UInt,PtrHostent,UInt,16)
	h_name      := ExtractInteger(hostent,0,false,4)
	h_aliases   := ExtractInteger(hostent,4,false,4)
	h_addrtype  := ExtractInteger(hostent,8,false,2)
	h_length    := ExtractInteger(hostent,10,false,2)
	h_addr_list := ExtractInteger(hostent,12,false,4)
	; // Retrieve official name
	VarSetCapacity(Name,64,0)
	DllCall("RtlMoveMemory",UInt,&Name,UInt,h_name,UInt,64)
	; // Retrieve Aliases
	VarSetCapacity(Aliases,12,0)
	DllCall("RtlMoveMemory", UInt, &Aliases, UInt, h_aliases, UInt, 12)
	Loop, 3
	{
		offset := ((A_Index-1)*4)
		PtrAlias%A_Index% := ExtractInteger(Aliases,offset,false,4)
		If (PtrAlias%A_Index% = 0)
			break
		VarSetCapacity(Alias%A_Index%,64,0)
		DllCall("RtlMoveMemory",UInt,&Alias%A_Index%,UInt,PtrAlias%A_Index%,Uint,64)
	}
	VarSetCapacity(AddressList,12,0)
	DllCall("RtlMoveMemory",UInt,&AddressList,UInt,h_addr_list,UInt,12)
	Loop, 3
	{
		offset := ((A_Index-1)*4)
		PtrAddress%A_Index% := ExtractInteger(AddressList,offset,false,4)
		If (PtrAddress%A_Index% =0)
			break
		VarSetCapacity(address%A_Index%,4,0)
		DllCall("RtlMoveMemory" ,UInt,&address%A_Index%,UInt,PtrAddress%A_Index%,Uint,4)
		i := A_Index
		Loop, 4
		{
		if Straddress%i%
			Straddress%i% := Straddress%i% "." ExtractInteger(address%i%,(A_Index-1 ),false,1)
		else
			Straddress%i% := ExtractInteger(address%i%,(A_Index-1 ),false,1)
		}
		Straddress0 = %i%
	}
	loop, %Straddress0% ; // put them together and return them
	{
		_this := Straddress%A_Index%
		if _this <>
			IPs = %IPs%%_this%
		if A_Index = %Straddress0%
			break
		IPs = %IPs%`n
	}
	return IPs
}
;----------------

ExtractInteger(ByRef sr, o = 0, is = false, s = 4)
{
	Loop %s%
		r += *(&sr + o + A_Index-1) << 8*(A_Index-1)
	If (!is OR s > 4 OR r < 0x80000000)
		return r
	return -(0xFFFFFFFF - r + 1)
}
;----------------

InsertInteger(i, ByRef d, o = 0, s = 4)
{
	Loop %s%
		DllCall("RtlFillMemory", "UInt", &d + o + A_Index-1, "UInt", 1, "UChar", i >> 8*(A_Index-1) & 0xFF)
}
;----------------

GuiEscape:
GuiClose:
GuiShowHide:
	GUI1 := WinExist()
	If DllCall( "IsWindowVisible", "UInt", GUI1)
		Gui, Hide
	else
		Gui, Show
return
;----------------

ExitSub:  ; // This subroutine is called automatically when the script exits for any reason.
	; // MSDN: "Any sockets open when WSACleanup is called are reset and automatically
	; // deallocated as if closesocket was called."
	DllCall("Ws2_32\WSACleanup")
	ExitApp
;----------------

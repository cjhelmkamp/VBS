'*********************************************************
' Name:	 	mapDrives.vbs
' Purpose:  Test if the computer is connected to the internet,
'			test if it is connected to the VPN, and then 
'			launch the mapped drives script 
'*********************************************************

On Error Resume Next			'Uncommnet to debug script

Dim ObjProgressMsg

strTargetInet = "8.8.4.4"		'Address used to test
strHqNet = "192.168.1.1"
strVpnName = "VPN"
strVpnNet = "192.168.99."
strWindowTitle = "Maped Drives, Connectivity test..." 

' The following makes the message window Global for the script
strComputer = "."
Set objWMIService = GetObject("Winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * From Win32_DesktopMonitor")
	For Each objItem in colItems
		intHorizontal = objItem.ScreenWidth
		intVertical = objItem.ScreenHeight
	Next
Set objExplorer = CreateObject _
	("InternetExplorer.Application")
	
Message "Checking for connectivity, please wait...", strWindowTitle, 0
WScript.Sleep 3000				'Sleep for 3 seconds

If Reachable(strTargetInet) Then
	Message "Connected to the Internet     ", strWindowTitle, 1
	WScript.Sleep 3000				'Sleep for 3 seconds
	Else
	Message "Not connected to internet     ", strWindowTitle, 2
	AlertMsg "Please connect to the internet to continue ", strWindowTitle
	WScript.quit
End If

If ((intVPNConnected(strVpnNet)) OR (Reachable(strHqNet))) Then
	Message "Connected to HQ Network       ", strWindowTitle, 1
	WScript.Sleep 3000				'Sleep for 3 seconds
	logonScript
	objExplorer.Quit
	WScript.Quit
	Else
	Message "Not Connected to VPN          ", strWindowTitle, 1
	WScript.Sleep 1500				'Sleep for 1.5 seconds
	Message "Attempting to connect to VPN, plese wait...", strWindowTitle, 1
	connectVPN(strVpnName)
	if intVPNConnected(strVpnNet) Then
		Message "Connected to VPN          ", strWindowTitle, 1
		WScript.Sleep 3000				'Sleep for 3 seconds
		logonScript
		objExplorer.Quit
		WScript.Quit
		Else
		Message "Failed to connect to VPN! ", strWindowTitle, 2
		AlertMsg "Please connect VPN or call the helpdesk for assistance", strWindowTitle
		WScript.quit
	End if
End If

Message "Done                     ", strWindowTitle, 2
WScript.quit

 '*********************************************************
 ' Name:	 intVPNConnected
 ' Purpose:  Test if the computer has VPN connected
 ' Input: 	 strVpnNet: IP of vpn network
 ' Returns:  If sussesful return True
 '           If failure returns False   
 '*********************************************************
Function intVPNConnected(strVpnNet)
	intVPNConnected = False				'Set false untill test is true
	sComputer = "." 
	Set oWMIService = GetObject("winmgmts:\\" & sComputer & "\root\CIMV2") 
	Set RouteTable = oWMIService.ExecQuery("select * from Win32_IP4RouteTable")
	For Each RouteEntry In RouteTable
		Caption = RouteEntry.Caption
		If InStr(Caption, strVpnNet) then
			intVPNConnected = true
		End if
	Next
End Function

Function connectVPN(strVpnName)
'*********************************************************
' Name:	 connectVPN
' Purpose:  Test if the computer has VPN connected
' Input: 	 strVpnName: Name of vpn to test for connection
' Returns:  If sussesful return True
'           If failure returns False   
'*********************************************************
	Set WSHShell = WScript.CreateObject("WScript.Shell")
	vbConnectionName = strVpnName
	vbConnectWith = "rasdial " &  """" & vbConnectionName & """" 
	WSHShell.run vbConnectWith,0,1		'command, visability, somethin
	WScript.Sleep 6000
	connectVPN = true
End Function

Function AlertMsg( strMessage, strWindowTitle )
'*********************************************************
' Title:	AlertMsg
' Author: 	Denis St-Pierre, Chris Helmkamp
' Purpose:  Displays a alert message box that the originating 
'			script can kill in both 2k and XP
' Input:	strMessage:		Message to show
'			strWindowTitle:	Title of window
'			If StrMessage is blank, take down previous 
'			alert message box
' Info:		Using 4096 in Msgbox below makes the alert message
'			float on top of things
' CAVEAT: 	You must have   Dim ObjProgressMsg   at the top of 
'			your script for this to work as described
'*********************************************************
    Set wshShell = CreateObject( "WScript.Shell" )
    strTEMP = wshShell.ExpandEnvironmentStrings( "%TEMP%" )
    If strMessage = "" Then
        ' Disable Error Checking in case objProgressMsg doesn't exists yet
        On Error Resume Next
        ' Kill ProgressMsg
        objProgressMsg.Terminate( )
        ' Re-enable Error Checking
        On Error Goto 0
        Exit Function
    End If
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    strTempVBS = strTEMP + "\" & "Message.vbs"     'Control File for reboot

    ' Create Message.vbs, True=overwrite
    Set objTempMessage = objFSO.CreateTextFile( strTempVBS, True )
    objTempMessage.WriteLine( "MsgBox""" & strMessage & """, 16+4096, """ & strWindowTitle & """" )
    objTempMessage.Close

    ' Disable Error Checking in case objProgressMsg doesn't exists yet
    On Error Resume Next
    ' Kills the Previous ProgressMsg
    objProgressMsg.Terminate( )
    ' Re-enable Error Checking
    On Error Goto 0

    ' Trigger objProgressMsg and keep an object on it
    Set objProgressMsg = WshShell.Exec( "%windir%\system32\wscript.exe " & strTempVBS )

    Set wshShell = Nothing
    Set objFSO   = Nothing
End Function

sub Message (strMessage, strWindowTitle, intMsgRun)
'*********************************************************
' Name:	 	message
' Purpose:  Display progress messages of the script
' Input: 	strMessage: Name of vpn to test for connection
'			strWindowTitle: Title to display
'			intMsgRun: 0 first, 1 normal, 2 exit
' INFO:		Used as a subroutine because there is no return
'*********************************************************
On Error Resume Next
startTextWrap = "<br><br><center><h3>"
endTextWrap = "</h3></center>"
if (intMsgRun = 0) Then
	
	objExplorer.Navigate "about:blank"   
	objExplorer.ToolBar = 0
	objExplorer.StatusBar = 0
	objExplorer.Left = (intHorizontal - 400) / 2
	objExplorer.Top = (intVertical - 200) / 2
	objExplorer.Width = 400
	objExplorer.Height = 200 
	objExplorer.Visible = 1             
	
	objExplorer.Document.Body.Style.Cursor = "wait"

	objExplorer.Document.Title = strWindowTitle
	objExplorer.Document.Body.InnerHTML = startTextWrap & strMessage & endTextWrap
	Elseif(intMsgRun = 1)Then
	objExplorer.Document.Body.InnerHTML = startTextWrap & strMessage & endTextWrap
	Else
	objExplorer.Document.Body.InnerHTML = startTextWrap & strMessage & endTextWrap
	objExplorer.Document.Body.Style.Cursor = "default"
	WScript.Sleep 3000				'Sleep for 3 seconds
	objExplorer.Quit
end if
end sub

sub logonScript
'*********************************************************
' Name:	 	logonScript
' Purpose:  Launch the logon script of the user
' INFO:		Used as a subroutine because there is no return.
'			Will need to be modified to call a script on a server.
'*********************************************************
	strLogonScript = "\\DOMAINNAME\NETLOGON\standard.vbs"
	Set wshShell = CreateObject("Wscript.shell")
	set oEnv = wshShell.Environment("PROCESS")
	' stop security warning about running from network
	oEnv("SEE_MASK_NOZONECHECKS") = 1
	'Make sure the file exists.  If exists run the file.  If not exists report an error.
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If objFSO.FileExists(strLogonScript) Then
		'Run the logon script
		wshShell.Run strLogonScript
	end if
	oEnv.Remove("SEE_MASK_NOZONECHECKS")		'remove suppression
end sub

Function Reachable(strComputer)
'==========================================================================
' The following function will test if a machine is reachable via a ping
' using WMI and the Win32_PingStatus Class
'==========================================================================
'On Error Resume Next
 
Dim wmiQuery, objWMIService, objPing, objStatus
  
wmiQuery = "Select * From Win32_PingStatus Where Address = '" & strComputer & "'"
  
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set objPing = objWMIService.ExecQuery(wmiQuery)
  
For Each objStatus in objPing
	If IsNull(objStatus.StatusCode) Or objStatus.Statuscode<>0 Then
		Reachable = False 'if computer is unreacable, return false
		Else
		Reachable = True 'if computer is reachable, return true
	End If
  Next
End Function
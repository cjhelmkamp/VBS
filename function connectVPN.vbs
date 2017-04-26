 '*********************************************************
 ' Name:	 connectVPN
 ' Purpose:  Test if the computer has VPN connected
 ' Input: 	 strVpnName: Name of vpn to test for connection
 ' Returns:  If sussesful return True
 '           If failure returns False   
 '*********************************************************
Function connectVPN(strVpnName)
	Set WSHShell = WScript.CreateObject("WScript.Shell")
	vbConnectionName = strVpnName
	'vbConnectionUser = "username"
	'vbConnectionPassword = "password"
	vbConnectWith = "rasdial " &  """" & vbConnectionName & """" '& """ """ & vbConnectionUser & """ """ & vbConnectionPassword & """"
	'vbConnectWith = "rasdial asdfasdf"
	'WScript.Sleep 20000
	'Wscript.Echo vbConnectWith
	WSHShell.run vbConnectWith,0,1
	WScript.Sleep 6000
	connectVPN = true
End Function
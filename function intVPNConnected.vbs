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
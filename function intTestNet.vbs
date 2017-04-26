 '*********************************************************
 ' Name:	 intTestNet
 ' Purpose:  Test if the computer has a connection with ping
 ' Input:    strTargetNet: ipaddress to ping
 ' Returns:  If sussesful return True
 '           If failure returns False   
 '*********************************************************
Function intTestNet (strTargetNet())
    Dim i 							'Loop counter
	Dim blnReply					'Target reply
	i = 0							'Start i at 0
	blnReply = False				'Set false untill test is true
	
	Do While i < 1 and not blnReply
		Set objShell = CreateObject("WScript.Shell")
		Set objExec = objShell.Exec("ping -n 2 -w 1000 " & strTargetNet)
		strPingResults = LCase(objExec.StdOut.ReadAll)
		'wscript.echo strPingResults
		If 0<((InStr(strPingResults, "reply from")) + (Not(InStr(strPingResults, "unreachable")))) Then
			blnReply = True			'Set flag True
			Else
			i=i+1					'Incriment the counter
			WScript.Sleep 9000		'Delay for network 
		End If
	Loop
	intTestNet = blnReply			'Set return value
End Function
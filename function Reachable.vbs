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
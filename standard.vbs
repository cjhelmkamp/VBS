On Error Resume Next

dim objNetwork, strDrive, objShell, objUNC 
dim strRemotePath, strDriveLetter, strNewName 

'Create the objects
set FSO   = CreateObject("Scripting.FileSystemObject")
'Set variable for the current Active Directory user
Set objSysInfo = CreateObject("ADSystemInfo") 
Set objUser = GetObject("LDAP://" & objSysInfo.UserName)
'Use the GetInfo method to initialize the local cache with attributes of the user account object
objUser.GetInfo
strUsername = objUser.sAMAccountName
strHomeDir = objUser.homedirectory
strWindowTitle = "Mapping Drives..."

'Allowes the messages to change
strComputer = "."
Set objWMIService = GetObject("Winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * From Win32_DesktopMonitor")
	For Each objItem in colItems
		intHorizontal = objItem.ScreenWidth
		intVertical = objItem.ScreenHeight
	Next
Set objExplorer = CreateObject _
	("InternetExplorer.Application")

Message "Remove and map drives starting...", strWindowTitle, 0

call cleardrives() 'Delete all mapped drives... 
call mapdrive("H:",strHomeDir,strUsername) 
call mapdrive("I:","\\server\folder","Ahima Share") 
call mapdrive("G:","\\server\folder","Department Share") 

sub cleardrives
	'Remove mapped network drives with exceptions
	Dim DriveCollection, Mapping
	Set objNetwork = CreateObject("Wscript.Network")
	Set DriveCollection = objNetwork.EnumNetworkDrives
	For Mapping = 0 To DriveCollection.Count - 1 Step 2
		If DriveCollection(Mapping) = "H:" Then        'Skip H: mapped in AD
		'ElseIf DriveCollection(Mapping) = "I:" Then   'Causes I: to be skipped 
		Message DriveCollection(Mapping) & " drive already mapped, skipping.", strWindowTitle, 1
		Else
			Message "Removing " & DriveCollection(Mapping), strWindowTitle, 1
			objNetwork.RemoveNetworkDrive DriveCollection(Mapping), True, True
		End If
	Next
end sub

function mapdrive(strDriveLetter,strRemotePath,strNewName) 
Err.clear
If FSO.DriveExists(strDriveLetter) = False Then
    Message "Mapping " & strDriveLetter & " " & strNewName, strWindowTitle, 1
    set objNetwork = createObject("WScript.Network")  
    objNetwork.MapNetworkDrive strDriveLetter, strRemotePath  
	 
    'Rename the Mapped Drive 
    set objShell = createObject("Shell.Application") 
    objShell.NameSpace(strDriveLetter).Self.name = strNewName 
end if
end function

'*********************************************************
 ' Name:	 message
 ' Purpose:  Display progress messages of the script
 ' Input: 	 strMessage: Name of vpn to test for connection
 '			 strWindowTitle: Title to display
 '			 intMsgRun: 0 first, 1 normal, 2 exit
 '*********************************************************
sub Message (strMessage, strWindowTitle, intMsgRun)
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
	WScript.Sleep 3000				'Sleep for 3 seconds
Elseif(intMsgRun = 1)Then
	objExplorer.Document.Body.InnerHTML = startTextWrap & strMessage & endTextWrap
	WScript.Sleep 3000				'Sleep for 3 seconds
Else
	objExplorer.Document.Body.InnerHTML = startTextWrap & strMessage & endTextWrap
	objExplorer.Document.Body.Style.Cursor = "default"
	WScript.Sleep 6000				'Sleep for 6 seconds
	objExplorer.Quit
end if
end sub

Message "Mapping drives complete", strWindowTitle, 1
Message "Exiting", strWindowTitle, 2
WScript.Quit
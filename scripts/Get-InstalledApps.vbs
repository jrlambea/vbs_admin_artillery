'
' Author Jose Ramón Lambea
'
' 140128 Script for query through wmi the installed applications with their
'		 product code, helpful for make a uninstall command with msiexec.
'
'		 Usage:
'
'             cscript.exe get-InstalledApps.vbs $host
'

Dim oApplication,colApplications,oService,objArgs,wmiServer

Set objArgs = Wscript.Arguments

If objArgs.Count <> 1 Then
	Wscript.Echo "Usage: cscript.exe get-InstalledApps.vbs $host"
	Wscript.Quit 5
End If

wmiServer = objArgs(0)

set oLocator = CreateObject("WbemScripting.SWbemLocator")

' The ConnectServer function has these arguments:
'	1 - Computer to query
'	2 - Namespace
'	3 - User name
'	4 - Password
set oService = oLocator.ConnectServer(wmiServer, "root\cimv2", "", "")


if Err = 0 then
	set colApplications = oService.InstancesOf("Win32_Product")

    oService.Security_.impersonationlevel = 3
    oService.Security_.Privileges.AddAsString "SeLoadDriverPrivilege"

else
	Wscript.Echo Err.Message
	Wscript.Quit 6

end if

Wscript.Echo "Retrieving installed applications from host: " & wmiServer & _
			 ", this may take a while."

for each oApplication in colApplications
	Wscript.Echo "Caption:" & oApplication.Caption
	Wscript.Echo "IdentifyingNumber:" & oApplication.IdentifyingNumber
	Wscript.Echo

Next

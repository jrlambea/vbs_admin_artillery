'
' Author Jose Ramón Lambea
'
' 140128 Script for uninstall applications through the Win32_Product class.
'
' Usage:
'             cscript.exe remove-InstalledApps.vbs $host $IdentifyingNumber
'
Dim oApplication,colApplications,oService,oArgs,wmiServer,bLoc,strId

Set oArgs = Wscript.Arguments
bLoc = False

If oArgs.Count <> 2 Then
     Wscript.Echo "Usage:" & vbcrlf & _
                  "cscript.exe get-InstalledApps.vbs $host $IdentifyingNumber"
     Wscript.Quit 5

End If

wmiServer = oArgs(0)
strId     = oArgs(1)

set oLocator = CreateObject("WbemScripting.SWbemLocator")

' The ConnectServer function has these arguments:
'        1 - Computer to query
'        2 - Namespace
'        3 - User name
'        4 - Password
set oService = oLocator.ConnectServer(wmiServer, "root\cimv2", "", "")

If Err = 0 Then
     set colApplications = oService.InstancesOf("Win32_Product")

     oService.Security_.impersonationlevel = 3
     oService.Security_.Privileges.AddAsString "SeLoadDriverPrivilege"

Else
     Wscript.Echo Err.Message
     Wscript.Quit 6

End If

Wscript.Echo "Retrieving installed applications from host: " & wmiServer & _
             ", this may take a while."

For Each oApplication In colApplications
    
        Select Case strId
            Case oApplication.Caption, oApplication.IdentifyingNumber, _
            oApplication.Name, oApplication.PackageCode
                bLoc = True
        End Select


    If bLoc Then
        Wscript.Echo "Application located, trying to uninstall " & oApplication.Caption & "..."
        oApplication.Uninstall

        If Not Err Then
            Wscript.Echo "Application " & oApplication.Caption & _
                         " uninstalled succesfully."
            Wscript.Quit 0
        Else
            Wscript.Echo "The program has not been removed." & vbcrlf & Err.Description
            Wscript.Quit 2
        End If

        Exit For

    End If

Next

If Not bLoc Then
    Wscript.Echo "The application isn't installed. No program has been removed."
    Wscript.Quit 1
End If

On Error Resume Next 

strPasswd = "Cyb3rPatr!0t$"
strComputer = "." 

Set objWMIService = GetObject("winmgmts:" _ 
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 

Set colItems = objWMIService.ExecQuery _ 
    ("Select * from Win32_UserAccount Where LocalAccount = True") 

For Each objItem in colItems 
    Do While True
        if objItem.Name = "Guest" then Exit Do ' Skip some account
        if objItem.Name = "Administrator" then Exit Do ' Skip some account
        if objItem.PasswordChangeable = False then Exit Do ' 

        objItem.SetPassword strPasswd
        objItem.SetInfo 

        Exit Do
    Loop
Next 

Wscript.Echo "Done."
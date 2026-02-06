Sub SAP_Logon()

Dim SapGui
Dim Applic
Dim connection
Dim session
Dim WSHShell

Shell "C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplgpad.exe", vbNormalFocus

Set WSHShell = CreateObject("WScript.Shell")

Do Until WSHShell.AppActivate("SAP Logon ")
    Application.Wait Now + TimeValue("0:00:01")
Loop

Set WSHShell = Nothing

Set SapGui = GetObject("SAPGUI")

Set Applic = SapGui.GetScriptingEngine

Set connection = Applic.OpenConnection("R3P - SAP ERP Production System", True)

Set session = connection.Children(0)

session.findById("wnd[0]").maximize

session.findById("wnd[0]/usr/txtRSYST-MANDT").Text = "050"
session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = Worksheets("Login").Range("D6").Value
session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = Worksheets("Login").Range("D8").Value
session.findById("wnd[0]/usr/txtRSYST-LANGU").Text = "PT"

session.findById("wnd[0]").sendVKey 0


Set session = Nothing

Set connection = Nothing

Set sap = Nothing

End Sub
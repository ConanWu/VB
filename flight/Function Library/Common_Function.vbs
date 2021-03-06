


'fValueVerify(strObj, strValue)
'fCheckObjectExist(strObj)
'Login(InputData)
'Logout(InputData)
'Login_old(InputData)
'Logout_old(InputData)
'Launch_Flight_App(InputData)
'Book_FlightOrder(InputData)


'----------re-defined qtp function------------
RegisterUserFunc "WpfWindow","CheckObjectExist","fCheckObjectExist",False
RegisterUserFunc "Window","CheckObjectExist","fCheckObjectExist",False
RegisterUserFunc "WinComboBox","ValueVerify","fValueVerify",False


'***********************************

'***********************************
Function fValueVerify(strObj, strValue)
	Set obj = strObj
	Dim var_ObjUIName : var_ObjUIName = obj.getROProperty("attached text")
	If Trim(obj.getROProperty("text")) = Trim(strValue) Then
		Reporter.ReportEvent micPass,"Verify value in "&var_ObjUIName&".","Value in "&var_ObjUIName&" is "&strValue&" successfully."
	Else
		Reporter.ReportEvent micFail,"Verify value in "&var_ObjUIName&".","Value of "&obj.getROProperty("text")&" in "&var_ObjUIName&" is not mathch with expected value "&strValue&"."
	End If
End Function

Function fCheckObjectExist(strObj)
	Dim ReturnValue : ReturnValue = 0
	Set obj = strObj
	gstrObjName = obj.GetROProperty("text")
	If obj.Exist(0) then
		Reporter.ReportEvent micPass,"Check for "&gstrObjName& " object exist","Object of "&gstrObjName& " is present."
	Else
		Reporter.ReportEvent micFail,"Check for "&gstrObjName& " object exist","Object of "&gstrObjName& " does not exist."
		ReturnValue = 1
	End If
	fCheckObjectExist = ReturnValue
End Function

Function Login(InputData)
	Dim ReturnValue : ReturnValue = 0
	Dim UserID :	UserID=Environment.Value("UserID")
	Dim Password :	Password=Environment.Value("Password")
	WpfWindow("Window_Flight").WpfEdit("WpfEdit_AgentName").type UserID
	WpfWindow("Window_Flight").WpfEdit("WpfEdit_Password").type Password
	WpfWindow("Window_Flight").WpfButton("WpfButton_OK").Click
	WpfWindow("Window_Flight").CheckObjectExist
	WpfWindow("Window_Flight").Activate
	Login = ReturnValue
End Function

Function Logout(InputData)
	Dim ReturnValue : ReturnValue = 0
	If WpfWindow("Window_Flight").Exist(2) Then WpfWindow("Window_Flight").Close
	If Not WpfWindow("Window_Flight").Exist(0) Then
		Reporter.ReportEvent micPass,"Verify logout successful","Logout successful"
	Else
		Reporter.ReportEvent micFail,"Verify logout successful","logout failed"
		ReturnValue = ReturnValue + 1
	End If
	Logout = ReturnValue
End Function

Function Login_old(InputData)
	Dim ReturnValue : ReturnValue = 0
	Dim UserID :	UserID=Environment.Value("UserID")
	Dim Password :	Password=Environment.Value("Password")
	Dialog("Dialog_Login").WinEdit("WinEdit_AgentName").type UserID
	Dialog("Dialog_Login").WinEdit("WinEdit_Password").type Password
	Dialog("Dialog_Login").WinButton("WinButton_OK").Click
	Window("Window_FlightReservation").CheckObjectExist
	Window("Window_FlightReservation").Activate
	Login_old = ReturnValue
End Function

Function Logout_old(InputData)
	Dim ReturnValue : ReturnValue = 0
	If Window("Window_FlightReservation").Exist(2) Then Window("Window_FlightReservation").Close
	If Not Window("Window_FlightReservation").Exist(0) Then
		Reporter.ReportEvent micPass,"Verify logout successful","Logout successful"
	Else
		Reporter.ReportEvent micFail,"Verify logout successful","logout failed"
		ReturnValue = ReturnValue + 1
	End If
	Logout_old = ReturnValue
End Function

Function Launch_Flight_App(InputData)
	Dim ReturnValue : ReturnValue = 0
	Dim flight_APPURL
	flight_APPURL=Environment.Value("flight_APPURL")
	SystemUtil.Run flight_APPURL
	Launch_Flight_App = ReturnValue
End Function

Function Book_FlightOrder(InputData)
	Window("Window_FlightReservation").Activate
	Window("Window_FlightReservation").ActiveX("ActiveX_MaskEdBox").Type "111113"
	Window("Window_FlightReservation").WinComboBox("WinComboBox_FlyFrom").Select "Denver"
	Window("Window_FlightReservation").WinComboBox("WinComboBoxFlyTo").Select "London"
	Window("Window_FlightReservation").WinButton("WinButton_FLIGHT").Click
	Window("Window_FlightReservation").Dialog("Dialog_FlightsTable").WinList("WinList_From").Select "20251   DEN   06:12 AM   LON   01:23 PM   AA     $112.20"
	Window("Window_FlightReservation").Dialog("Dialog_FlightsTable").WinButton("WinButton_OK").Click
	Window("Window_FlightReservation").WinEdit("WinEdit_Name").Set "Conan"
	Window("Window_FlightReservation").WinEdit("WinEdit_Tickets").Set "2"
	Window("Window_FlightReservation").WinButton("WinButton_InsertOrder").Click
End Function

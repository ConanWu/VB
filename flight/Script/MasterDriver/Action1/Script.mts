


Environment.LoadFromFile(Split(Environment.Value("TestDir"),"Script")(0)&"FrameWork\EnvironmentVaribale.xml")
Set fso = CreateObject("Scripting.FileSystemObject")
DriverPath = Split(Environment.Value("TestDir"),"MasterDriver")(0)&"Master_Driver.xls"
If fso.FileExists(DriverPath) Then
	datatable.ImportSheet DriverPath,"Action1","Action1"
	DriverRow = datatable.GetSheet("Action1").GetRowCount
	If DriverRow < 0 Then
		Reporter.ReportEvent micFail,"To find the master driver data","Failed to find the master driver data"
		exitrun
	End If
Else
	Reporter.ReportEvent micFail,"Verify  the master driver file exist","Master driver are not exist, please check driver name"
	exitrun
End If


FW_ExecuteScript

exittest







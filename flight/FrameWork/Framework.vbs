

' FW_ExecuteScript()
' FW_Parameter()
' FW_GetCurrentRow()


'*****************************************************************************************************************
' Function Name   	: FW_ExecuteScript()
' Purpose          	: Main Driver Function 
' Input Parameters	: None
' Output Parameter	: 
' Date of creation 	: 09/03/2014
' Author Name     	: Conan
' Change History   	: v1.0
' Calling Method   	:
'*****************************************************************************************************************
Function FW_ExecuteScript()
	
	For Counter = 1 to DriverRow
		datatable.GetSheet("Action1").SetCurrentRow(Counter)
		If Ucase(Trim(datatable.Value("ToBeExcute","Action1"))) = "YES" Then
			Set g_Variable = CreateObject("Scripting.dictionary")
			Call FW_LoadVariables()														'load global variables'

			datatable.ImportSheet g_Variable("dataFilePath"),"Global","Global"			'import datafile to global datatable
			call EnterNode(g_Variable("Test_Case"),"Below are the case content.")
			g_Variable.Add "g_row",FW_GetCurrentRow()
			If g_Variable("g_row") < 0 Then
				Reporter.ReportEvent micFail,"Find test data","Failed to find the test data."
				Exit Function
			End If

			FinalResult = 0
			Datatable.SetCurrentRow(g_Variable("g_row"))
			Do while (Ucase(datatable.Value("Test_Case")) = Ucase(g_Variable("Test_Case")))
				g_Variable("Step_No") = datatable.Value("Step_No")
				g_Variable("g_step") = Datatable.Value("Step")
				call EnterNode(g_Variable("Step_No")&"--"&g_Variable("g_step"),"Below are the step content.")
				If FinalResult = 0 Then
					If Instr(g_Variable("g_step"),"InputData") > 0 Then
						Set InputData = FW_Parameter()
					End If
					FinalResult = Eval(g_Variable("g_step"))
					Call FW_Log(FinalResult,g_OutputKey,"")
				End If
				Reporter.ReportEvent micDone,g_Variable("g_step")&" is done.","Test case "&g_Variable("Test_Case")
				call ExitNode()
				g_Variable("g_row") = g_Variable("g_row") + 1
				datatable.SetCurrentRow(g_Variable("g_row"))
			Loop
			If FinalResult = 0 Then
				Reporter.ReportEvent micPass,g_Variable("Test_Case")&" is Passed.","Test case "&g_Variable("Test_Case")
				datatable.GetSheet("Action1").SetCurrentRow(Counter)
				'Datatable.Value("ToBeExcute","Action1") = "PASS"
			Else
				Reporter.ReportEvent micFail,g_Variable("Test_Case")&" is Failed.","Test case "&g_Variable("Test_Case")
			End If
			call ExitNode()
			
		End If

	Next
	Call FW_Log(FinalResult,g_OutputKey,"FinalResult")
	datatable.ExportSheet DriverPath,"Action1"							'export master datafile to global datatable
	'datatable.ExportSheet g_Variable("outputFilePath"),"Summary"		'export result report to summary sheet
	'datatable.ExportSheet g_Variable("outputFilePath"),"TestStepResult"	'export detail result to teststepresult

	Set g_Variable = Nothing
End Function


Sub FW_LoadVariables()
	g_Variable.Add "TestStartTime",now()
	g_Variable.Add "TestStartTimer",timer()
	g_Variable.Add "TestNo",datatable.value("TestNo","Action1")
	g_Variable.Add "Test_Case",datatable.value("Test_Case","Action1")
	g_Variable.Add "Failed_Reason",""
	g_Variable.Add "DataFile",datatable.value("DataFile","Action1")
	g_Variable.Add "dataFilePath",Split(Environment.Value("TestDir"),"Script")(0)&"DataFile\"&g_Variable("DataFile")
	g_Variable.Add "templatePath",Split(Environment.Value("TestDir"),"Script")(0)&"Function Library\"&"TestResultTemplate.xls"
	g_Variable.Add "outputFilePath",Split(Environment.Value("TestDir"),"Script")(0)&"log\Summary\"&"ReportSummary_"&Replace(Date(),"/","_")&".xls"
End Sub

'*****************************************************************************************************************
' Function Name   	: FW_Parameter()
' Purpose          	: Add the parameters into dictionary object 
' Input Parameters	: None
' Output Parameter	: 
' Date of creation 	: 09/04/2014
' Author Name     	: Conan
' Change History   	: v1.0
' Calling Method   	:
'*****************************************************************************************************************
Function FW_Parameter()
	Set dicObj = CreateObject("Scripting.Dictionary")
	ColumnNumber = datatable.GetSheet("Global").GetParameterCount
	For i = 1 to ColumnNumber
		ColumnName = Datatable.GetSheet("Global").GetParameter(i).Name
		dicObj.Add ColumnName,Datatable.Value(ColumnName)
	Next
	Set FW_Parameter = dicObj
	Set dicObj = nothing
	 
End Function

'*****************************************************************************************************************
' Function Name   	: FW_GetCurrentRow()
' Purpose          	: Find the test case in data file and reture the line row number 
' Input Parameters	: None
' Output Parameter	: 
' Date of creation 	: 09/04/2014
' Author Name     	: Conan
' Change History   	: v1.0
' Calling Method   	:
'*****************************************************************************************************************
Function FW_GetCurrentRow()
	RowCount = datatable.GetSheet("Global").GetRowCount
	FinalRow = -1
	For CurrentRow = 1 to RowCount
		datatable.SetCurrentRow CurrentRow
		If Ucase(datatable.Value("Test_Case")) = Ucase(g_Variable("Test_Case")) Then
			FinalRow = CurrentRow
			Exit For
		End If
	Next
	FW_GetCurrentRow = FinalRow
End Function

'*****************************************************************************************************************
' Function Name   	: FW_Log()
' Purpose          	: log the test result in excel
' Input Parameters	: None
' Output Parameter	: 
' Date of creation 	: 04/06/2016
' Author Name     	: Conan
' Change History   	: v1.0
' Calling Method   	:
'*****************************************************************************************************************
Function FW_Log(strResult,OutputKey,strType)
	Set fso = CreateObject("Scripting.FileSystemObject")
	DataTable.AddSheet("Summary")
	DataTable.AddSheet("TestStepResult")
	If fso.FileExists(g_Variable("outputFilePath")) Then
		datatable.ImportSheet g_Variable("outputFilePath"),"Summary","Summary"
		datatable.ImportSheet g_Variable("outputFilePath"),"TestStepResult","TestStepResult"
	Else
		datatable.ImportSheet g_Variable("templatePath"),"Summary","Summary"
		datatable.ImportSheet g_Variable("templatePath"),"TestStepResult","TestStepResult"
	End If
	If strType = "FinalResult" Then
		If strResult > 0 Then
			strDesc = "Fail"
		Else
			strDesc = "Pass"
		End If
		Sum_RowCount = datatable.GetSheet("Summary").GetRowCount
		datatable.GetSheet("Summary").SetCurrentRow Sum_RowCount+1

		Reporter.ReportEvent strResult,g_Variable("Test_Case"),strDesc

		datatable.value("TestNo","Summary") = g_Variable("TestNo")
		datatable.value("Test_Case","Summary") = g_Variable("Test_Case")
		datatable.value("TestStartTime","Summary") = g_Variable("TestStartTime")
		datatable.value("TestEndTime","Summary") = now()
		datatable.value("TestTotalExecuteTime","Summary") = CInt(timer()-g_Variable("TestStartTimer"))&"second"
		datatable.value("Test_Status","Summary") = strDesc
		DataTable.value("Path_Result_File","Summary") = Environment("ResultDir")
		datatable.value("Failed_Reason","Summary") = g_Variable("Failed_Reason")

	Else
		If strResult > 0 Then
			strDesc = "Fail"
		Else
			strDesc = "Pass"
		End If
		Sum_RowCount = datatable.GetSheet("TestStepResult").GetRowCount
		datatable.GetSheet("TestStepResult").SetCurrentRow Sum_RowCount+1
'		Reporter.ReportEvent strResult,
		datatable.Value("TestNo","TestStepResult")=g_Variable("TestNo")
		datatable.value("Test_Case","TestStepResult") = g_Variable("Test_Case")
		datatable.Value("Step_No","TestStepResult")=g_Variable("Step_No")
		datatable.Value("Step_Name","TestStepResult")=g_Variable("g_step")
		datatable.Value("Test_Status","TestStepResult")=strDesc
		
		a=1
		'datatable.GetSheet("TestStepResult").AddParameter "",""
	End If

	datatable.ExportSheet g_Variable("outputFilePath"),"Summary"
	datatable.ExportSheet g_Variable("outputFilePath"),"TestStepResult"
End Function



Public Function EnterNode(ByRef NodeName, ByRef NodeContent)
  Set dicMetaDescription = CreateObject("Scripting.Dictionary")
  dicMetaDescription("Status") = MicDone
  dicMetaDescription("PlainTextNodeName") = NodeName
  dicMetaDescription("StepHtmlInfo") = NodeContent
  dicMetaDescription("DllIconIndex") = 210
  dicMetaDescription("DllIconSelIndex") = 210
  dicMetaDescription("DllPAth") = Environment.Value("ProductDir")&"\bin\ContextManager.dll"
  intContext = Reporter.LogEvent("User", dicMetaDescription, Reporter.GetContext)
  Reporter.SetContext intContext
End Function

Public Function ExitNode()
   Reporter.UnSetContext
End Function

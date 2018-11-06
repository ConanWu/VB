SystemUtil.Run("iexplore.exe")
Call SystemUtil.Run("iexplore.exe","http://www.baidu.com")
SystemUtil.CloseProcessByName("iexplore.exe")

SystemUtil.Run("notepad.exe")
SYstemUtil.CloseProcessByWndTitle("Untitled - Notepad")



'close process from task manager
strSQL = "Select * From Win32_Process Where Name = 'iexplore.exe'"
strSQL1 = "select * from Win32_Process"
Set oWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set ProcColl = oWMIService.ExecQuery(strSQL)

For Each oElem in ProcColl
	if oElem.name = "iexplore.exe" Then	oElem.Terminate
Next

Set oWMIService = Nothing
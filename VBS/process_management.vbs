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


'resource 利用GetObject("WinMgmts:")获取系统信息
'用WMI对象列出系统所有进程: 
'----Instance.vbs---- 
'程序代码 
Dim WMI,objs
Set WMI = GetObject("WinMgmts:")
Set objs = WMI.InstancesOf("Win32_Process") 
For Each obj In objs 
	Enum1 = Enum1 + obj.Description + Chr(13) + Chr(10) 
Next
msgbox Enum1


'获得物理内存的容量: 
'-----physicalMemory.vbs----- 
'' 程序代码 
strComputer = "." 

Set wbemServices = GetObject("winmgmts:\\" & strComputer) 
Set wbemObjectSet = wbemServices.InstancesOf("Win32_LogicalMemoryConfiguration") 

For Each wbemObject In wbemObjectSet 
	WScript.Echo "物理内存 (MB): " & CInt(wbemObject.TotalPhysicalMemory/1024)
Next 


'取得系统所有服务及运行状态 
'----service.vbs---- 
'' 程序代码 
Set ServiceSet = GetObject("winmgmts:").InstancesOf("Win32_Service")
Dim s,infor
infor=""
for each s in ServiceSet
	infor=infor+s.Description+" ==> "+s.State+chr(13)+chr(10)
next
msgbox infor

'用WMI对象列出系统所有进程: 
'----Instance.vbs----
'' 程序代码 
Dim WMI,objs
Set WMI = GetObject("WinMgmts:")
Set objs = WMI.InstancesOf("Win32_Process") 
For Each obj In objs 
	Enum1 = Enum1 + obj.Description + Chr(13) + Chr(10) 
Next
msgbox Enum1


'获得物理内存的容量: 
'-----physicalMemory.vbs-----
'' 程序代码 
strComputer = "." 

Set wbemServices = GetObject("winmgmts:\\" & strComputer) 
Set wbemObjectSet = wbemServices.InstancesOf("Win32_LogicalMemoryConfiguration") 

For Each wbemObject In wbemObjectSet 
	WScript.Echo "物理内存 (MB): " & CInt(wbemObject.TotalPhysicalMemory/1024)
Next 


'取得系统所有服务及运行状态 
'----service.vbs----
'' 程序代码 
Set ServiceSet = GetObject("winmgmts:").InstancesOf("Win32_Service")
Dim s,infor
infor=""
for each s in ServiceSet
	infor=infor+s.Description+" ==> "+s.State+chr(13)+chr(10)
next
msgbox infor


'CPU的序列号: 
'---CPUID.vbs--- 
'' 程序代码 
Dim cpuInfo
cpuInfo = ""
set moc = GetObject("Winmgmts:").InstancesOf("Win32_Processor")
for each mo in moc
	cpuInfo = CStr(mo.ProcessorId)
	msgbox "CPU SerialNumber is : " & cpuInfo
next


'硬盘型号: 
'---HDID.vbs---
'' 程序代码 
Dim HDid,moc
set moc =GetObject("Winmgmts:").InstancesOf("Win32_DiskDrive")
for each mo in moc
	HDid = mo.Model
	msgbox "硬盘型号为:" & HDid
next


'网卡MAC物理地址: 
'---MACAddress.vbs---
'' 程序代码 
Dim mc
set mc=GetObject("Winmgmts:").InstancesOf("Win32_NetworkAdapterConfiguration")
for each mo in mc
if mo.IPEnabled=true then
	msgbox "网卡MAC地址是: " & mo.MacAddress
	exit for
end if
next


'测试你的显卡: 
'' 程序代码 
On Error Resume Next
Dim ye
Dim yexj00 
set yexj00=GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_VideoController") 
for each ye in yexj00 
	msgbox " 型 号: " & ye.VideoProcessor & vbCrLf & " 厂 商: " & ye.AdapterCompatibility & vbCrLf & " 名 称: " & ye.Name & vbCrLf & " 状 态: " & ye.Status & vbCrLf & " 显 存: " & (ye.AdapterRAM\1024000) & "MB" & vbCrLf & "驱 动 (dll): " & ye.InstalledDisplayDrivers & vbCrLf & "驱 动 (inf): " & ye.infFilename & vbCrLf & " 版 本: " & ye.DriverVersion
next

'补充一段结束所有qq进程代码：

strComputer="." 
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2") 
Set colProcessList=objWMIService.ExecQuery ("select * from Win32_Process where Name='QQ.exe' ") 
For Each objProcess in colProcessList 
	objProcess.Terminate() 
next
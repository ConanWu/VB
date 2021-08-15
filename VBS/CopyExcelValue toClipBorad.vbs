
'create text file to clipboard value
Set oFile = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\1.txt",8,true)

'create htmlfile to get clipboard value
Set oClipBoardTxt = CreateObject("htmlfile")

'create excel to open and copy the excel value
Set oExcel = CreateObject("Excel.Application")
oExcel.Visible = false
Set oWb = oExcel.Workbooks.Open("C:\1.xls")
Set oSheet = oWb.Sheets("Sheet1")
oSheet.Range("C1:C100").Copy

'get the clipboard value and write to text file
clipboardtxt = oClipBoardTxt.ParentWindown.ClipBoardData.getData("text")
oFile.Write(clipboardtxt)

oExcel.DisplayAlerts = false
oWb.Close false
oExcel.Quit
Set oExcel = nothing
Set Owb = nothing
Set oSheet = nothing
oClipBoardTxt.close
oFile.close
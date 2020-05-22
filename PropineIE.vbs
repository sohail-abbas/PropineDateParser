Dim ie
Dim objExcel
Dim objWorkbook
Dim objSheet
Dim intCol
Dim strTestData
Dim rowCount
Dim outCol
Dim rloop
Dim sOutPutPath
Dim resultCol
Dim intPassCount
Dim intFailCount
Dim sTestExcelPath

strUrl = "https://vast-dawn-73245.herokuapp.com/"
set WshShell = WScript.CreateObject("WScript.Shell") 
sOutPutPath = "C:\Personal\Propine\Output"
strExcelPath = "C:\Personal\Propine\TestSheet.xlsx"
strReportTemplate = "C:\Personal\Propine\Template\reportTemplate.html"
strFileName = "C:\Personal\Propine\test.txt"
strFileName1 = "C:\Personal\Propine\test2.txt"

intCol = 3
outCol = 5
resultCol = 6 
intPassCount = 0
intFailCount = 0 


'-------------------------------------------Master Scritp Call---------------------------------
fnCleanUp()
fnOpenExcel strExcelPath
fnRowColCount()
sTestExcelPath = sOutPutPath & "\TestResult_" & Year(Now) &month(now) &day(now) &hour(now) &minute(now) &second(now) & ".xlsx"
fnCopyFile strExcelPath, sTestExcelPath
rloop = 2
fnLaunchIE strUrl

	'This segment is to perform sanity test, if sanity fails test will stop else will continue. 
	fnReadExcelCell rloop,intCol
	fnEnterValueinWeb strTestData	
	fnReadNotepad strFileName
	fnReadtext strFileName1
	fnWriteValueinExcelCell rloop, outCol, strnewValue
	WScript.Sleep 5000

	'This loop will iterate for all test scenario in Input Sheet. 
	fnCloseExcel()
	for rloop = 2 to rowCount
			fnOpenExcel sTestExcelPath
			fnReadExcelCell rloop,intCol
			fnEnterValueinWeb strTestData
			fnReadNotepad strFileName
			fnReadtext strFileName1
			fnWriteValueinExcelCell rloop, outCol, strnewValue
			WScript.Sleep 5000
			fnCloseExcel()
	next
	
fnCreateHTMLReport()
msgbox "Execution Completed"

'--------------------------------------------------Functions Library---------------------------------------

'-------------------------------
'Function Name 	:	fnCleanUp
'Variables		:	None
'Description	:	This function is to cleanup all the temporary file and kill any open excel or internet explorer to have clean run.
'-------------------------------
function fnCleanUp()
	fnDeleteNotepad strFileName
	fnDeleteNotepad strFileName1
	fnKillProcess "EXCEL.EXE"
	fnKillProcess "IEXPLORE.EXE"
end function

'-------------------------------
'Function Name 	:	fnLaunchIE
'Variables		:	sUrl (Url to be navigating after opening Internet Explorer)
'Description	:	This function is open internet explorer and navigate to url provided. 
'-------------------------------
function fnLaunchIE(sUrl)
	On Error Resume Next
	Set ie = CreateObject("InternetExplorer.Application")
	ie.Navigate sUrl 
	ie.Visible = True
	WScript.Sleep 5000
	WshShell.AppActivate "IE"
end function

'-------------------------------
'Function Name 	:	fnkeypress
'Variables		:	strValue (key value provided)
'Description	:	This function is to press any key event. 
'-------------------------------
function fnkeypress(strValue)
	WshShell.SendKeys strValue
end function

'-------------------------------
'Function Name 	:	fnOpenExcel
'Variables		:	sPath (Path of excel sheet)
'Description	:	This function is open excel workbook.
'-------------------------------
function fnOpenExcel(sPath)
	Set objExcel = CreateObject("Excel.Application")
	Set objWorkbook = objExcel.Workbooks.Open(sPath)
	Set objSheet=objWorkbook.Worksheets("Input") 

	objExcel.Visible = False
end function

'-------------------------------
'Function Name 	:	fnReadExcelCell
'Variables		:	strRow,strColumn  (Row and column number to be provided as input)
'Description	:	This function is read any cell value
'-------------------------------
function fnReadExcelCell(strRow, strColumn)
	strTestData = objSheet.Cells(strRow, strColumn).Value
end function

'-------------------------------
'Function Name 	:	fnWriteValueinExcelCell
'Variables		:	strRow, strColumn, strValue  
'Description	:	This function is to write value in the exact cell in which user wants. 
'-------------------------------
function fnWriteValueinExcelCell(strRow, strColumn, strValue)
	objSheet.Cells(strRow, strColumn).Value = strValue
	objWorkbook.Save()
end function

'-------------------------------
'Function Name 	:	fnRowColCount
'Variables		:	None
'Description	:	This function is to read row count of the excel 
'-------------------------------
function fnRowColCount()
	rowCount = objSheet.usedrange.rows.count  
end function

'-------------------------------
'Function Name 	:	fnCloseExcel
'Variables		:	None
'Description	:	This function is to close Excel and free up memory 
'-------------------------------
function fnCloseExcel()
	objWorkbook.Close
	objExcel.Quit

	Set objSheet = Nothing
	Set objWorkbook = Nothing
	Set objExcel = Nothing
end function

'-------------------------------
'Function Name 	:	fnEnterValueinWeb
'Variables		:	strTestData (Value provided as input)
'Description	:	This function is to enter value in Date text box of the date parser webpage.
'-------------------------------
function fnEnterValueinWeb(strTestData)	
	ie.document.All.Item("date").Value = strTestData
	
	WshShell.AppActivate "IE" 

	for i = 1 to 8 
		fnkeypress "{TAB}"
	next
	fnkeypress "{ENTER}"
	WshShell.AppActivate "IE" 
	WScript.Sleep 5000
	fnReadPage()
end function

'-------------------------------
'Function Name 	:	fnReadPage
'Variables		:	None
'Description	:	This function is to read inner text of the webpage 
'-------------------------------
function fnReadPage()
	Dim ie 
	Dim objShell 
	Dim objWindow 
	Dim objItem 

	Set objShell = CreateObject("Shell.Application")
	Set objWindow = objShell.Windows()
	For Each objItem In objWindow
		If Instr(1, Lcase(objItem.FullName), "iexplore.exe", 1) <> 0 Then
			Set ie = objItem
		End If
	Next 

	fnOpenNotepad strFileName, ie.Document.body.innertext
end function

'-------------------------------
'Function Name 	:	fnOpenNotepad
'Variables		:	strFileName, strTextValue
'Description	:	This function is to open a notepad and save the page inner text.  
'-------------------------------
function fnOpenNotepad(strFileName, strTextValue)
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set oFile = FSO.CreateTextFile(strFileName)
	oFile.WriteLine strTextValue 
	oFile.Close
	Set fso = Nothing
	Set oFile = Nothing    	
end function

'-------------------------------
'Function Name 	:	fnReadNotepad
'Variables		:	strFileName (file path)
'Description	:	This function is to open and read notepad data 
'-------------------------------
function fnReadNotepad(strFileName)
	On Error Resume Next
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set txtStream = fso.OpenTextFile(strFileName)

	Do While Not txtStream.AtEndOfStream
		'msgbox txtStream.ReadLine
		fnOpenNotepad strFileName1, txtStream.ReadLine
		if Instr(1,txtStream.ReadLine, "GMT", vbTextCompare) > 0 then 
			strnewValue = txtStream.ReadLine
		end if		
	Loop
	txtStream.Close
	Set fso = Nothing
	Set txtStream = Nothing  
end function

'-------------------------------
'Function Name 	:	fnDeleteNotepad
'Variables		:	strFileName (file path)
'Description	:	This function is to delete the file if present. 
'-------------------------------
function fnDeleteNotepad(strFileName)
	Set fso = CreateObject("Scripting.FileSystemObject")
	If fso.FileExists(strFileName) Then
		FSO.DeleteFile(strFileName)
	end if 
	Set fso = Nothing   	
end function

'-------------------------------
'Function Name 	:	fnReadtext
'Variables		:	strFileName (file path)
'Description	:	This function is to read line from the file
'-------------------------------
function fnReadtext(strFileName)
	'On Error Resume Next
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set txtStream = fso.OpenTextFile(strFileName)

	Do While Not txtStream.AtEndOfStream
			strnewValue = txtStream.ReadLine
	Loop
	txtStream.Close
	Set fso = Nothing
	Set txtStream = Nothing  
end function

'-------------------------------
'Function Name 	:	fnWriteLineToFile
'Variables		:	strPath, strData
'Description	:	This function is to open a file and write line at the end of file 
'-------------------------------
function fnWriteLineToFile(strPath, strData)
	Set oFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(strPath,8,true)
	oFileToWrite.WriteLine(strData)
	oFileToWrite.Close
	Set oFileToWrite = Nothing
end function 

'-------------------------------
'Function Name 	:	fnReplaceText
'Variables		:	strPath, strFromText, strToText
'Description	:	This function is to open a file and replae any text. 
'-------------------------------
function fnReplaceText(strPath, strFromText, strToText)
	Set objFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(strPath)

	strText = objFile.ReadAll
	objFile.Close
	Set objFile = nothing
	strNewText = Replace(strText, strFromText, strToText)
	
	Set oFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(strPath, 2, true)
	oFile.WriteLine strNewText 
	oFile.Close
	Set oFile = nothing
end function 

'-------------------------------
'Function Name 	:	fnKillProcess
'Variables		:	strProcessName
'Description	:	This function is to kill any process in order to free the memory
'-------------------------------
function fnKillProcess(strProcessName)
	On Error Resume Next
	Dim Process 	
	For Each Process In GetObject("winmgmts:").ExecQuery("Select Name from Win32_Process Where Name = '"&strProcessName&"'") 
		Process.Terminate 
	Next 
end function


'-------------------------------
'Function Name 	:	fnCopyFile
'Variables		:	strSource, strTarget
'Description	:	This function is to copy file from one location to another 
'-------------------------------
function fnCopyFile(strSource, strTarget)
	CreateObject("Scripting.FileSystemObject").CopyFile strSource, strTarget, true
end function

'-------------------------------
'Function Name 	:	fnCreateHTMLReport
'Variables		:	None
'Description	:	This function is to create HTML Report.
'-------------------------------
function fnCreateHTMLReport()
	fnOpenExcel sTestExcelPath
	fnRowColCount()
	strReport = sOutPutPath & "\TestResult_" & Year(Now) &month(now) &day(now) &hour(now) &minute(now) &second(now) & ".html"
	fnCopyFile strReportTemplate, strReport
	for sloop = 2 to rowCount
		fnReadExcelCell sloop,resultCol
		if strTestData = "PASS" then 
			intPassCount = intPassCount + 1
		else	
			intFailCount = intFailCount + 1
		end if 
	next 
	fnWriteLineToFile strReport, ""
	for i = 2 to rowCount
		for j = 1 to resultCol
			fnReadExcelCell i,j
			if strTestData = "PASS" then 
				fnWriteLineToFile strReport, "<td style=xxxx1background-color:#008000;text-align:centerxxxx1><font color=xxxx1#ffffffxxxx1><b>PASS</td>"
			elseif strTestData = "FAIL" then 
				fnWriteLineToFile strReport, "<td style=xxxx1background-color:#ff0000;text-align:centerxxxx1><font color=xxxx1#ffffffxxxx1><b>FAIL</td>"
			else	
				fnWriteLineToFile strReport, "<td>"&strTestData&"</td>"	
			end if			
		next 
		fnWriteLineToFile strReport, "</tr><tr>"
	next 	
	fnWriteLineToFile strReport, "</tr></table><pre> </pre></body></html>"	
	fnCloseExcel()
	fnCleanUp()
	fnReplaceText strReport, "xxxx1", """"
	iTotScenario = intPassCount + intFailCount
	iPwidth = round((intPassCount/iTotScenario)*100*0.8,2)
	iFwidth = round((intFailCount/iTotScenario)*100*0.8,2)
	iPpercent = round((intPassCount/iTotScenario)*100,2)
	iFpercent = round((intFailCount/iTotScenario)*100,2)
	fnReplaceText strReport, "pwidth", iPwidth
	fnReplaceText strReport, "fwidth", iFwidth
	fnReplaceText strReport, "pPercent", "Count: " & intPassCount & "; Percentage: "& iPpercent
	fnReplaceText strReport, "fPercent", "Count: " & intFailCount & "; Percentage: "& iFpercent	
	fnLaunchIE strReport
end function 

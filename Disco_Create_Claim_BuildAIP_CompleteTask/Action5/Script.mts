Dim ExcelObj, WorkBook
'Click & Navigate to Disco Serach Screen
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=.*Search","xpath:=//DIV[@id='discoMenuBar']/SPAN[3]").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=.*Search","xpath:=//DIV[@id='discoMenuBar']/SPAN[3]").Click
wait 5

'Wait Property for the Search Administration & Blank Search
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("html id:=workRequestLabelId","xpath:=//DIV[@id='workRequestLabelId']").CheckProperty "innertext", "Work Request #"
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("html id:=searchButtonId","xpath:=//DIV[@id='searchButtonId']").Highlight

If Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Search Criteria","xpath:=//DIV[@id='searchCriteriaLabelId']").Exist Then
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Search Criteria","xpath:=//DIV[@id='searchCriteriaLabelId']").Highlight
	Wait 2
	XlPath  = "C:\IDBPM_UFT\Test_Data\Test_Data.xls"
	Set ExcelObj = CreateObject("Excel.Application")
 	ExcelObj.Visible = True
 	Set WorkBook = ExcelObj.Workbooks.Open(XlPath)
 	TaskCwpId = WorkBook.worksheets(1).Cells(2,18).Value 
 	'msgbox TaskCwpId
 	'msgbox WorkBook.worksheets(1).Cells(2,18).Value 
  
	Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("html id:=cwpSearchId","xpath:=//INPUT[@id='cwpSearchId']").Set TaskCwpId
	Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("html id:=searchButtonId","xpath:=//DIV[@id='searchButtonId']").CheckProperty "html id", "searchButtonId"
	Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("html id:=searchButtonId","xpath:=//DIV[@id='searchButtonId']").Click		
	If Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=.*is closed in Vantive and therefore AIPs .*","xpath:=//DIV[@role='alert']").Exist Then
		Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=.*is closed in Vantive and therefore AIPs .*","xpath:=//DIV[@role='alert']").Highlight
		Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=.*is closed in Vantive and therefore AIPs .*","xpath:=//DIV[@role='alert']").FireEvent "onmouseover"
		Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=.*is closed in Vantive and therefore AIPs .*","xpath:=//DIV[@role='alert']").Click
		Reporter.ReportEvent micStatus, "AIP Search Result's with more than 30Day's Closure" , "Search on AIP - PASS"
		Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("html id:=totalMRRLabelId","xpath:=//DIV[@id='totalMRRLabelId']").Highlight
		Else
			Reporter.ReportEvent micPass, "CWP Search Result's for Disconnect Request" , "Search on CWP's working as Expected - PASS"
			Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("html id:=actualETCLabelId","xpath:=//DIV[@id='actualETCLabelId']").Highlight
			Wait 3
			Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=.*Dashboard","xpath:=//DIV[@id='discoMenuBar']/SPAN[1]").Click
	End If
End If

ExcelObj.Application.Quit
Set ExcelObj=nothing
Browser("title:=BPM Console").CloseAllTabs

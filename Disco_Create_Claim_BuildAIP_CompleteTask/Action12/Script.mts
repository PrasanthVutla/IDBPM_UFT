'Variables
Dim CurDate, ReqReceiveDate, LodReceiveDate, CustWantDate, DelSiteDate, CurMonth, CurDay, CurYear, modifiedDay, modifiedYear, WantSiteDate, TaskCwpId, ExcelObj ,MySheet
CurDate = Date
CurMonth = MonthName(Month(Now()),1)
CurYear = Year(Now())
CurDay = Day(Now())
CustWantDate = dateadd("d",30,CurDate)
DelSiteDate = dateadd("d",30,CurDate)
modifiedDay = MonthName(Month(CustWantDate),1)
modifiedYear = Year(CustWantDate)
ReqReceiveDate = CurMonth&" " & CurDay&" " & CurYear 
LodReceiveDate = CurMonth&" " & CurDay&" " & CurYear
WantSiteDate = modifiedDay&" " & CurDay&" "& modifiedYear

'msgbox CurMonth&" " & CurDay&" " & CurYear
'msgbox modifiedDay&" " & CurDay&" "& modifiedYear

'Double Click on Disconnect Task
With Browser("title:=BPM Console").Page("title:=BPM Console")
	.WebTable("Class Name:=WebTable","class:=v-table-table","text:=Disconnects Task.*").WebElement("Class Name:=WebElement","class:=v-table-cell-wrapper","innertext:=Disconnects Task").FireEvent "onmouseover"
	.WebTable("Class Name:=WebTable","class:=v-table-table","text:=Disconnects Task.*").WebElement("Class Name:=WebElement","class:=v-table-cell-wrapper","innertext:=Disconnects Task").Highlight
	If Browser("title:=BPM Console").Page("title:=BPM Console").WebTable("Class Name:=WebTable","class:=v-table-table","text:=Disconnects Task.*").WebElement("Class Name:=WebElement","class:=v-table-cell-wrapper","innertext:=Disconnects Task").Exist Then
		Reporter.ReportEvent micPass, "Disconnect Task Available", "Task Available"
		.WebTable("Class Name:=WebTable","class:=v-table-table","text:=Disconnects Task.*").WebElement("Class Name:=WebElement","class:=v-table-cell-wrapper","innertext:=Disconnects Task").FireEvent"ondblclick"
		wait 8
		Else
			Reporter.ReportEvent micFail, "Disconnect Task Not-Available", "Task Available - Failed"
	End If
End With

'Verify & Claim Disconnect Task
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","xpath:=//DIV[@id='taskMenuBar']/SPAN[1]","innertext:=.*Claim").CheckProperty "innertext", "Claim"
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","xpath:=//DIV[@id='taskMenuBar']/SPAN[1]","innertext:=.*Claim").Click
Wait 5
If Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("xpath:=//DIV[@role='alert']","innertext:=Info: Sucessfully claimed.*").Exist Then
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("xpath:=//DIV[@role='alert']","innertext:=Info: Sucessfully claimed.*").Highlight
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("xpath:=//DIV[@role='alert']","innertext:=Info: Sucessfully claimed.*").FireEvent "onmouseover"
	Reporter.ReportEvent micPass, "Disconnect Task Sucessfully claimed", "Claim Successful"	
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("class:=v-nativebutton-caption","innertext:=Disco").FireEvent "onmouseover"
	Wait 3
End If
Browser("title:=BPM Console").Page("title:=BPM Console").WebTable("Class Name:=WebTable","class:=v-table-table","text:=Disconnects Task.*").WebElement("Class Name:=WebElement","class:=v-table-cell-wrapper","innertext:=Disconnects Task").FireEvent "onmouseover"
Browser("title:=BPM Console").Page("title:=BPM Console").WebTable("Class Name:=WebTable","class:=v-table-table","text:=Disconnects Task.*").WebElement("Class Name:=WebElement","class:=v-table-cell-wrapper","innertext:=Disconnects Task").FireEvent"ondblclick"
If Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-loading-indicator second v-loading-indicator-delay").Exist Then
	wait 15
End If

'Fill the Dico Accordian details
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-menubar-menuitem","innertext:=.*Revoke").FireEvent "onmouseover"
'Add Validate for Complete Button - Disable/Enable
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-menubar-menuitem v-menubar-menuitem-disabled","innertext:=.*Complete").Highlight

'Click & Fill Critical Dates Accordian
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-captiontext","outertext:=ASSOCIATED RESPONSIBILITY").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-captiontext","outertext:=CRITICAL DATES").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-captiontext","outertext:=CRITICAL DATES").Click
Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("Class Name:=WebEdit","xpath:=//*[@id='requestReceivedDate']/INPUT").Set ReqReceiveDate
Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("Class Name:=WebEdit","xpath:=//*[@id='lodReceivedDate']/INPUT").Set LodReceiveDate
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-menubar-menuitem","innertext:=.*Save").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-menubar-menuitem","innertext:=.*Save").Click
Wait 6

'Click & Fill Reporting Attributes Accordian
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-captiontext","outertext:=REPORTING ATTRIBUTES").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-captiontext","outertext:=REPORTING ATTRIBUTES").Click
Browser("title:=BPM Console").Page("title:=BPM Console").WebList("Class Name:=WebList","html id:=reasonCode").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebList("Class Name:=WebList","html id:=reasonCode").WebButton("xpath:=//DIV[@id='reasonCode']/DIV[@role='button']").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebList("Class Name:=WebList","html id:=reasonCode").WebButton("xpath:=//DIV[@id='reasonCode']/DIV[@role='button']").Click
Browser("title:=BPM Console").Page("title:=BPM Console").WebTable("Class Name:=WebTable","innertext:=.*internal.*").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","innerhtml:=Next").Click
Browser("title:=BPM Console").Page("title:=BPM Console").WebTable("Class Name:=WebTable","innertext:=Service IssuesSh.*").FireEvent "onmouseover"
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","xpath:=//TR/TD[@role='listitem' and normalize-space()='TBD']").Click

Browser("title:=BPM Console").Page("title:=BPM Console").WebList("Class Name:=WebList","html id:=subReasonCode").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebList("Class Name:=WebList","html id:=subReasonCode").WebButton("xpath:=//DIV[@id='subReasonCode']/DIV[@role='button']").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebList("Class Name:=WebList","html id:=subReasonCode").WebButton("xpath:=//DIV[@id='subReasonCode']/DIV[@role='button']").Click
Browser("title:=BPM Console").Page("title:=BPM Console").WebTable("Class Name:=WebTable","innertext:=.*Clean Up.*").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","innerhtml:=Next").Click
Wait 2
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","innerhtml:=Next").Click
Wait 2
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","innerhtml:=Next").Click
Wait 2
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","innerhtml:=Next").Click
Browser("title:=BPM Console").Page("title:=BPM Console").WebTable("Class Name:=WebTable","innertext:=HAN/Inkra.*").FireEvent "onmouseover"
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","xpath:=//TR/TD[@role='listitem' and normalize-space()='TBD']").Click

Browser("title:=BPM Console").Page("title:=BPM Console").WebList("Class Name:=WebList","html id:=siteDeleteType").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebList("Class Name:=WebList","html id:=siteDeleteType").WebButton("xpath:=//DIV[@id='siteDeleteType']/DIV[@role='button']").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebList("Class Name:=WebList","html id:=siteDeleteType").WebButton("xpath:=//DIV[@id='siteDeleteType']/DIV[@role='button']").Click
Browser("title:=BPM Console").Page("title:=BPM Console").WebTable("Class Name:=WebTable","innertext:=Full.*").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","innerhtml:=Full").Click

Browser("title:=BPM Console").Page("title:=BPM Console").WebList("Class Name:=WebList","html id:=accountDeleteType").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebList("Class Name:=WebList","html id:=accountDeleteType").WebButton("xpath:=//DIV[@id='accountDeleteType']/DIV[@role='button']").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebList("Class Name:=WebList","html id:=accountDeleteType").WebButton("xpath:=//DIV[@id='accountDeleteType']/DIV[@role='button']").Click
Browser("title:=BPM Console").Page("title:=BPM Console").WebTable("Class Name:=WebTable","innertext:=Full.*").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","innerhtml:=Partial").Click

Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-menubar-menuitem","innertext:=.*Save").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-menubar-menuitem","innertext:=.*Save").Click
Wait 6

'Click & Fill CWP Attributes
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-captiontext","outertext:=CWP ATTRIBUTES").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-captiontext","outertext:=CWP ATTRIBUTES").Click
Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("Class Name:=WebEdit","xpath:=//*[@id='customerWantDate']/INPUT").Set WantSiteDate
Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("Class Name:=WebEdit","xpath:=//*[@id='deleteSiteDate']/INPUT").Set WantSiteDate

Browser("title:=BPM Console").Page("title:=BPM Console").WebList("Class Name:=WebList","html id:=assignTo").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebList("Class Name:=WebList","html id:=assignTo").WebButton("xpath:=//DIV[@id='assignTo']/DIV[@role='button']").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebList("Class Name:=WebList","html id:=assignTo").WebButton("xpath:=//DIV[@id='assignTo']/DIV[@role='button']").Click
Browser("title:=BPM Console").Page("title:=BPM Console").WebTable("Class Name:=WebTable","innertext:=.*Asia On.*").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","innerhtml:=Asia OnDemand").Click
TaskCwpId = Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("xpath:=//DIV[normalize-space()='CWP Id']/INPUT").GetROProperty("Value")

Set ExcelObj = CreateObject("Excel.Application")
'ExcelObj.Application.Visible = true
wait 2
ExcelObj.Workbooks.Open "C:\IDBPM_UFT\Test_Data\Test_Data.xls" 
ExcelObj.Application.Visible = true
set MySheet = ExcelObj.ActiveWorkbook.Worksheets("Test_Data")
MySheet.cells(2,18).value = TaskCwpId
ExcelObj.ActiveWorkbook.Save

Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-menubar-menuitem","innertext:=.*Save").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-menubar-menuitem","innertext:=.*Save").Click
Wait 10

'Click AIP's/Service ID's Review
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-captiontext","outertext:=AIPS/SERVICE IDS REVIEW").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-captiontext","outertext:=AIPS/SERVICE IDS REVIEW").Click
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("xpath:=//DIV[@id='buildAipGridButtonMenuBar']/SPAN","innertext:=Build Aip Grid","visible:=True").CheckProperty "innertext", "Build Aip Grid"

ExcelObj.Application.Quit
Set ExcelObj=nothing

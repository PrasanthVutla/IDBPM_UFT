'Variables
Dim CurDate, ReqReceiveDate, LodReceiveDate, CustWantDate, DelSiteDate, CurMonth, CurDay, CurYear, modifiedDay, modifiedYear, WantSiteDate
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
	.WebTable("Class Name:=WebTable","class:=v-table-table","text:=Disconnects Task.*").WebElement("Class Name:=WebElement","class:=v-table-cell-wrapper","innertext:=Disconnects Task").FireEvent"ondblclick"
End With

'Verify & Claim Disconnect Task
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-captiontext","innertext:=Candidate Task-.*").WaitProperty "innerhtml", "Candidate Task.*"

Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-menubar-menuitem","innertext:=.*Claim").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-menubar-menuitem-caption","innertext:=.*Claim").Click
wait(10)

Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("class:=v-nativebutton-caption","innertext:=Disco").FireEvent "onmouseover"
Browser("title:=BPM Console").Page("title:=BPM Console").WebTable("Class Name:=WebTable","class:=v-table-table","text:=Disconnects Task.*").WebElement("Class Name:=WebElement","class:=v-table-cell-wrapper","innertext:=Disconnects Task").FireEvent"ondblclick"

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
Wait(6)

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
If Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-loading-indicator second v-loading-indicator-delay").Exist Then
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-loading-indicator second v-loading-indicator-delay").Highlight
	wait 10
End If

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

Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-menubar-menuitem","innertext:=.*Save").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-menubar-menuitem","innertext:=.*Save").Click
Wait 10

'Click AIP's/Service ID's Review
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-captiontext","outertext:=AIPS/SERVICE IDS REVIEW").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-captiontext","outertext:=AIPS/SERVICE IDS REVIEW").Click


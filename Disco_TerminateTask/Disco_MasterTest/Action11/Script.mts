'Variables
Dim curDate
curDate = Date

'Create New Disconnect Request
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("class:=v-button v-widget icon-only v-button-icon-only small v-button-small","html id:=createNewDisconnectRequestButton").Click

'Sync for New Disconnect Form
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("class:=v-window-header","innertext:=New Disconnect Request").WaitProperty "innertext", "New Disconnect Request", 10000

'Filling New Disco Form
'Customer Id
Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("Class Name:=WebEdit","html id:=customeridId").Set DataTable.Value("CustId")

'Disconnect Date
Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("Class Name:=WebEdit","class:=v-textfield v-datefield-textfield").Set curDate

'Disconnect Reason
Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("Class Name:=WebEdit","html id:=disconnectreasonId").Set DataTable.Value("DiscoReason")

'Disconnect Type
Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("Class Name:=WebEdit","html id:=disconnecttypeId").Set DataTable.Value("DiscoType")

'Technical Contact Name
Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("Class Name:=WebEdit","html id:=technicalcontactnameId").Set DataTable.Value("TechContName")

'WorkRequest Number
Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("Class Name:=WebEdit","html id:=workrequestnumberId").Set DataTable.Value("WR_No")

Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Start","xpath:=//DIV[@id='startMenuItemid']").WaitProperty "visible", True, 10000
Wait(5)

If Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Error: Disconnect request is already created.*","xpath:=//DIV[@role='alert'][1]").Exist Then
	
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Error: Disconnect request is already created.*","xpath:=//DIV[@role='alert'][1]").Highlight
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Error: Disconnect request is already created.*","xpath:=//DIV[@role='alert'][1]").Click
	
	Else
		Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Start","class:=v-menubar-menuitem").Click
		Reporter.ReportEvent micFail, "Disconnect Created - Duplicate", "Negative Scenario - Fail"
End If

Reporter.ReportEvent micPass, "Disconnect request is already created", "Negative Scenario - PASS"
wait 5
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("name:=close button","xpath:=//DIV[@role='button'][2]").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("name:=close button","xpath:=//DIV[@role='button'][2]").Click

'Click Disco Plugin
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("innertext:=.*Disco","xpath:=//DIV/BUTTON[normalize-space()='Disco']").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("innertext:=.*Disco","xpath:=//DIV/BUTTON[normalize-space()='Disco']").Click

'Sync for Disco Plugin
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("class:=v-label v-widget h1 v-label-h1 v-label-undef-w","xpath:=//DIV[normalize-space()='Disco']/DIV[1]").WaitProperty "innertext", "Disco", 10000
wait(10)
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("class:=v-label v-widget h1 v-label-h1 v-label-undef-w","xpath:=//DIV[normalize-space()='Disco']/DIV[1]").Highlight


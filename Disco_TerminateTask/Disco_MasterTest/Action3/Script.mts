'Variables
Dim curDate
curDate = Date

'Create New Disconnect Request
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("class:=v-button v-widget icon-only v-button-icon-only small v-button-small","html id:=createNewDisconnectRequestButton").Click
If Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-loading-indicator second v-loading-indicator-delay").Exist Then
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-loading-indicator second v-loading-indicator-delay").Highlight
	wait 10
End If

'Sync for New Disconnect Form
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("class:=v-window-header","innertext:=New Disconnect Request").WaitProperty "innertext", "New Disconnect Request", 10000

'Filling New Disco Form
'Customer Id
Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("Class Name:=WebEdit","html id:=customeridId").Set DataTable.Value("CustId")

'Customer Name
Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("Class Name:=WebEdit","html id:=customernameId").Set DataTable.Value("CustName")

'Customer Number
Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("Class Name:=WebEdit","html id:=customernumberId").Set DataTable.Value("CustNo")

'Disconnect Date
Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("Class Name:=WebEdit","class:=v-textfield v-datefield-textfield").Set curDate

'Disconnect Reason
Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("Class Name:=WebEdit","html id:=disconnectreasonId").Set DataTable.Value("DiscoReason")

'Disconnect Type
Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("Class Name:=WebEdit","html id:=disconnecttypeId").Set DataTable.Value("DiscoType")

'Main Contact Email
Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("Class Name:=WebEdit","html id:=maincontactemailId").Set DataTable.Value("MainContEmail")

'Main Contact Name
Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("Class Name:=WebEdit","html id:=maincontactnameId").Set DataTable.Value("MainContName")

'Submitted By
Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("Class Name:=WebEdit","html id:=submittedbyId").Set DataTable.Value("SubmittedBy")

'Technical Contact Email
Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("Class Name:=WebEdit","html id:=technicalcontactemailId").Set DataTable.Value("TechContEmail")

'Technical Contact Name
Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("Class Name:=WebEdit","html id:=technicalcontactnameId").Set DataTable.Value("TechContName")

'WorkRequest Number
Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("Class Name:=WebEdit","html id:=workrequestnumberId").Set DataTable.Value("WR_No")

Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Start","class:=v-menubar-menuitem").WaitProperty "visible", True, 10000
If Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-loading-indicator second v-loading-indicator-delay").Exist Then
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-loading-indicator second v-loading-indicator-delay").Highlight
	wait 15
End If
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Start","class:=v-menubar-menuitem").Click


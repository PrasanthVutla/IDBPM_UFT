'Click Disco Plugin
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("innertext:=.*Disco","xpath:=//DIV/BUTTON[normalize-space()='Disco']").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("innertext:=.*Disco","xpath:=//DIV/BUTTON[normalize-space()='Disco']").Click
If Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-loading-indicator second v-loading-indicator-delay").Exist Then
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-loading-indicator second v-loading-indicator-delay").Highlight
	wait 20
End If

'Sync for Disco Plugin
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("class:=v-label v-widget h1 v-label-h1 v-label-undef-w","xpath:=//DIV[normalize-space()='Disco']/DIV[1]").WaitProperty "innertext", "Disco", 10000
wait(10)
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("class:=v-label v-widget h1 v-label-h1 v-label-undef-w","xpath:=//DIV[normalize-space()='Disco']/DIV[1]").Highlight

'Click & Navigate to Configuration Tab - Disco
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=.*Configuration","xpath:=//DIV[@id='discoMenuBar']/SPAN[4]").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=.*Configuration","xpath:=//DIV[@id='discoMenuBar']/SPAN[4]").Click

'Wait Property for the Config Tab
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Email Templates","xpath:=//TBODY/TR/TD[@role='tab' and normalize-space()='Email Templates']").WaitProperty "innertext", "Email Templates", 5000
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Email Templates","class:=v-captiontext","xpath:=//DIV[@role='tablist']/TABLE[@role='presentation']/TBODY/TR/TD[@role='tab']/DIV[normalize-space()='Email Templates']/DIV[1]/DIV[1]").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Email Templates","class:=v-captiontext","xpath:=//DIV[@role='tablist']/TABLE[@role='presentation']/TBODY/TR/TD[@role='tab']/DIV[normalize-space()='Email Templates']/DIV[1]/DIV[1]").Click
Wait 5

'Add New Email Template
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("class:=v-tabsheet-content","innertext:=.*Email Template.*").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("html id:=addDiscoEmailTemplateButtonId","xpath:=//DIV[@id='addDiscoEmailTemplateButtonId']").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("html id:=addDiscoEmailTemplateButtonId","xpath:=//DIV[@id='addDiscoEmailTemplateButtonId']").Click
Wait 5

'Write the Email Template
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Email Editor.*","xpath:=//DIV[@role='dialog']").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("html id:=templateNameId","xpath:=//INPUT[@id='templateNameId']").Set DataTable.Value("DiscoReason")
Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("html id:=subjectId","xpath:=//INPUT[@id='subjectId']").Set DataTable.Value("TechContName")
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("title:=Insert Ordered List","xpath:=//DIV[@id='bodyId']//DIV[@role='button'][13]").Image("Class Name:=Image","xpath:=//DIV[@role='button'][13]/IMG[1]").Highlight
'Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("html id:=saveButtonId","xpath:=//DIV[@id='saveButtonId']").Highlight
'Browser("title:=BPM Console").Page("title:=BPM Console").Image("Class Name:=Image","xpath:=//DIV[@role='button'][13]/IMG[1]").Submit
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("html id:=saveButtonId","xpath:=//DIV[@id='saveButtonId']").Highlight
'Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("html id:=saveButtonId","xpath:=//DIV[@id='saveButtonId']").Click
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("name:=close button","xpath:=//DIV[@role='dialog'][1]/DIV[1]/DIV[1]/DIV[@role='button'][2]").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("name:=close button","xpath:=//DIV[@role='dialog'][1]/DIV[1]/DIV[1]/DIV[@role='button'][2]").Click
Wait 5
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Confirm.*","xpath:=//DIV[@id='confirmdialog-window']").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("html id:=confirmdialog-ok-button","xpath:=//DIV[@id='confirmdialog-ok-button']").Click
Wait 5

'Validate the Created Template & Get Deleted
If Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=AutoM Testing","xpath:=//TD/DIV[normalize-space()='AutoM Testing']").Exist Then
	Reporter.ReportEvent micPass, "Template Created - Successfully", "AutoM Testing Email Template Created"
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=AutoM Testing","xpath:=//TD/DIV[normalize-space()='AutoM Testing']").Highlight
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=AutoM Testing","xpath:=//TD/DIV[normalize-space()='AutoM Testing']").Click
	Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("html id:=deleteDiscoEmailTemplateButton","xpath:=//DIV[@id='deleteDiscoEmailTemplateButton']").Highlight
	Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("html id:=deleteDiscoEmailTemplateButton","xpath:=//DIV[@id='deleteDiscoEmailTemplateButton']").Click
	wait 4
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Delete Record.*","xpath:=//DIV[@role='dialog']").Highlight
	Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("innertext:=Ok","xpath:=//DIV[@role='button' and normalize-space()='Ok']").Highlight
	Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("innertext:=Ok","xpath:=//DIV[@role='button' and normalize-space()='Ok']").Click
	Reporter.ReportEvent micStatus, "Template Created", "Created Template - Deleted Successful"
	
	'Negative Case for Template not Created
	Else 
		Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=.*Please fill in.*","xpath:=//DIV[2]/DIV[@role='alert'][1]").Highlight
		Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=.*Please fill in.*","xpath:=//DIV[2]/DIV[@role='alert'][1]").FireEvent "onmouseover"
		wait 3		
		Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=.*Please fill in.*","xpath:=//DIV[2]/DIV[@role='alert'][1]").Click
		Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("name:=close button","xpath:=//DIV[@role='dialog'][1]/DIV[1]/DIV[1]/DIV[@role='button'][2]").Highlight
		Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("name:=close button","xpath:=//DIV[@role='dialog'][1]/DIV[1]/DIV[1]/DIV[@role='button'][2]").Click
		Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("innertext:=CANCEL","xpath:=//DIV[@id='confirmdialog-cancel-button']").Highlight
		Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("innertext:=CANCEL","xpath:=//DIV[@id='confirmdialog-cancel-button']").Click
		Reporter.ReportEvent micFail, "Template Not Created - Verify the Manadatory Fields", "Please Verify the Manadatory Fields & Confirm"	
End If
'Navigate back to Disco Plugin
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("innertext:=.*Disco","xpath:=//DIV/BUTTON[normalize-space()='Disco']").Click

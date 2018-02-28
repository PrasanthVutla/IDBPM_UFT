'Click Disco Plugin
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("innertext:=.*Disco","xpath:=//DIV/BUTTON[normalize-space()='Disco']").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("innertext:=.*Disco","xpath:=//DIV/BUTTON[normalize-space()='Disco']").Click
If Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-loading-indicator second v-loading-indicator-delay").Exist Then
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-loading-indicator second v-loading-indicator-delay").Highlight
	wait 15
End If

'Sync for Disco Plugin
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("class:=v-label v-widget h1 v-label-h1 v-label-undef-w","xpath:=//DIV[normalize-space()='Disco']/DIV[1]").WaitProperty "innertext", "Disco", 10000
wait(10)
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("class:=v-label v-widget h1 v-label-h1 v-label-undef-w","xpath:=//DIV[normalize-space()='Disco']/DIV[1]").Highlight

'Apply Filter for Un-Assigned Tasks
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("xpath:=//DIV[@id='filterButton']","name:=filterButton").FireEvent"onmouseover"
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("xpath:=//DIV[@id='filterButton']","name:=filterButton").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("xpath:=//DIV[@id='filterButton']","name:=filterButton").Click
Wait (5)
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("class:=v-nativebutton-caption","innertext:=Disco").FireEvent "onmouseover"
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","html id:=filter-field-select").Click
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","html id:=filter-field-select").WebButton("Class Name:=WebButton","class:=v-filterselect-button").Click
Wait(5)
Browser("title:=BPM Console").Page("title:=BPM Console").WebTable("Class Name:=WebTable","innertext:=Assignee.*","name:=WebTable").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebTable("Class Name:=WebTable","innertext:=Assignee.*","name:=WebTable").FireEvent "onmouseover"
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","xpath:=//TR/TD[@role='listitem' and normalize-space()='Assignee']","innertext:=Assignee").Click
Wait(5)

'Apply Filter for Operator is-empty
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("html id:=filter-operator-select","Class Name:=WebElement").Click
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("html id:=filter-operator-select","Class Name:=WebElement").WebButton("Class Name:=WebButton","class:=v-filterselect-button").Click
Wait(5)
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("class:=v-filterselect-suggestpopup v-filterselect-suggestpopup-form-property","Class Name:=WebElement").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("class:=v-filterselect-suggestpopup v-filterselect-suggestpopup-form-property","Class Name:=WebElement").FireEvent "onmouseover"
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","innertext:=is empty","xpath:=//TR/TD[@role='listitem' and normalize-space()='is empty']").Click
Wait(5)

'Click on Add Button once it is enabled
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-menubar-menuitem","innertext:=Add").Click
Wait(10)

'Verify the Filter is Applied or Not
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","xpath:=//DIV[normalize-space()='Assignee is empty']","innerhtml:=Assignee is empty.*").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","xpath:=//DIV[normalize-space()='Assignee is empty']","innerhtml:=Assignee is empty.*").WaitProperty "innertext", "Assignee is empty"

'Select Multiple Tasks
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-csslayout v-layout v-widget layout-panel v-csslayou.*", "outertext:=.*Task Name Priority.*").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","xpath:=//DIV[@id='nullTable']//DIV/TABLE[1]/TBODY[1]/TR[2]/TD[1]").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","xpath:=//DIV[@id='nullTable']//DIV/TABLE[1]/TBODY[1]/TR[2]/TD[1]").Click

'Implement SendKeys for selecting multiple Rows
Dim mySendKeys
set mySendKeys = CreateObject("WScript.shell")
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","xpath:=//DIV[@id='nullTable']//DIV/TABLE[1]/TBODY[1]/TR[2]/TD[1]").Click
'mySendKeys.SendKeys("{+}") 
mySendKeys.SendKeys("+{DOWN 3}")
wait 5

Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("html id:=claimTasksId","xpath:=//DIV[@id='claimTasksId']").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("html id:=claimTasksId","xpath:=//DIV[@id='claimTasksId']").Click

If Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Info: Successfully claimed the tasks","xpath:=//DIV[@role='alert']").Exist Then
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Info: Successfully claimed the tasks","xpath:=//DIV[@role='alert']").Highlight
	Wait 3
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Info: Successfully claimed the tasks","xpath:=//DIV[@role='alert']").FireEvent "onmouseover"
	Reporter.ReportEvent micPass, "Claim Multiple Disconnect Tasks" , "Disconnect Task Multi claim is working as Expected - Pass"
	
	Else
		'Clear the Filter & Report as Failure
		Reporter.ReportEvent micFail, "Claim Multiple Disconnect Tasks - Not Working" , "Disconnect Task Claim is not working as Expected - Fail"
		Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("xpath:=//DIV[@id='filterEditorView']/DIV[3]/DIV[1]/DIV[2]/DIV[1]").Highlight
		Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("xpath:=//DIV[@id='filterEditorView']/DIV[3]/DIV[1]/DIV[2]/DIV[1]").Click
		Wait (5)
		Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("xpath:=//DIV[@id='filterButton']","name:=filterButton").FireEvent"onmouseover"
		Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("xpath:=//DIV[@id='filterButton']","name:=filterButton").Highlight
		Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("xpath:=//DIV[@id='filterButton']","name:=filterButton").Click
		
End If

'Capture Bitmap for TestResults
Browser("title:=BPM Console").Page("title:=BPM Console").CaptureBitmap "ClaimTasks.png"
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("html id:=Refresh","xpath:=//DIV[@id='Refresh']").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("html id:=Refresh","xpath:=//DIV[@id='Refresh']").Click
Wait 5
Browser("title:=BPM Console").Page("title:=BPM Console").CaptureBitmap "ClaimAfterRefresh.png"

'Clear the Filter
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("xpath:=//DIV[@id='filterEditorView']/DIV[3]/DIV[1]/DIV[2]/DIV[1]").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("xpath:=//DIV[@id='filterEditorView']/DIV[3]/DIV[1]/DIV[2]/DIV[1]").Click
Wait (5)
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("xpath:=//DIV[@id='filterButton']","name:=filterButton").FireEvent"onmouseover"
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("xpath:=//DIV[@id='filterButton']","name:=filterButton").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("xpath:=//DIV[@id='filterButton']","name:=filterButton").Click
Reporter.ReportEvent micDone, "Claim Multiple Tasks - Successful" , "Working as Expected - Pass"



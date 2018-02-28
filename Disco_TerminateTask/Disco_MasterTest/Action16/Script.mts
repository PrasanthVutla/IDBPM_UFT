'Click Disco Plugin
datatable.Import "C:\IDBPM_UFT\Test_Data\Test_Data.xls"
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("innertext:=.*Disco","xpath:=//DIV/BUTTON[normalize-space()='Disco']").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("innertext:=.*Disco","xpath:=//DIV/BUTTON[normalize-space()='Disco']").Click
If Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-loading-indicator second v-loading-indicator-delay").Exist Then
	wait 25
End If

'Sync for Disco Plugin
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("class:=v-label v-widget h1 v-label-h1 v-label-undef-w","xpath:=//DIV[normalize-space()='Disco']/DIV[1]").WaitProperty "innertext", "Disco", 10000
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("class:=v-label v-widget h1 v-label-h1 v-label-undef-w","xpath:=//DIV[normalize-space()='Disco']/DIV[1]").Highlight

'Apply Filter for Work Request Number
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("Class:=v-button v-widget icon-only v-button-icon-only","name:=filterButton").FireEvent"onmouseover"
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("Class:=v-button v-widget icon-only v-button-icon-only","name:=filterButton").Click
Wait 5
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("class:=v-nativebutton-caption","innertext:=Disco").FireEvent "onmouseover"
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","html id:=filter-field-select").Click
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","html id:=filter-field-select").WebButton("Class Name:=WebButton","class:=v-filterselect-button").Click
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","innerhtml:=Next").Click
Wait 5
Browser("title:=BPM Console").Page("title:=BPM Console").WebTable("Class Name:=WebTable","innertext:=Submitted ByTask NameWorkRequest Number","name:=WebTable").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebTable("Class Name:=WebTable","innertext:=Submitted ByTask NameWorkRequest Number","name:=WebTable").FireEvent "onmouseover"
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=gwt-MenuItem","innertext:=WorkRequest Number").Click
Wait 5

'Apply Filter for Operator
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("html id:=filter-operator-select","Class Name:=WebElement").Click
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("html id:=filter-operator-select","Class Name:=WebElement").WebButton("Class Name:=WebButton","class:=v-filterselect-button").Click
Wait 5
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("class:=v-filterselect-suggestpopup v-filterselect-suggestpopup-form-property","Class Name:=WebElement").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("class:=v-filterselect-suggestpopup v-filterselect-suggestpopup-form-property","Class Name:=WebElement").FireEvent "onmouseover"
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=gwt-MenuItem","innertext:=equals").Click
Wait 5

'Apply Filter for Value
Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("html id:=filter-value-input","Class Name:=WebEdit").Set DataTable.Value("TerminateWR_No")

'Click on Add Button once it is enabled
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-menubar-menuitem","innertext:=Add").Click
Wait 10

'Verify the Filter is Applied or Not
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-label v-widget label v-label-label v-label-undef-w","innerhtml:=WorkRequest Number equals.*").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-label v-widget label v-label-label v-label-undef-w","innerhtml:=WorkRequest Number equals.*").WaitProperty "innerhtml", "WorkRequest Number equals "&WR_NO


If Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innerhtml:=Set Priority","xpath:=//TD/DIV[normalize-space()='Set Priority']").Exist Then
	
	'Double Click on Set Priority Task
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innerhtml:=Set Priority","xpath:=//TD/DIV[normalize-space()='Set Priority']").FireEvent "onmouseover"
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innerhtml:=Set Priority","xpath:=//TD/DIV[normalize-space()='Set Priority']").Highlight
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innerhtml:=Set Priority","xpath:=//TD/DIV[normalize-space()='Set Priority']").FireEvent"ondblclick"

	'Verify & Claim Set Priority Task
	If Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-captiontext","innertext:=Candidate Task-.*").Exist Then
		Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-captiontext","innertext:=Candidate Task-.*").WaitProperty "innerhtml", "Candidate Task.*"
		Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-menubar-menuitem","innertext:=.*Claim").Highlight
		Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-menubar-menuitem-caption","innertext:=.*Claim").Click
		wait 10
		Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("class:=v-nativebutton-caption","innertext:=Disco").FireEvent "onmouseover"
		Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innerhtml:=Set Priority","xpath:=//TD/DIV[normalize-space()='Set Priority']").FireEvent"ondblclick"
		
		'Open the Task
		Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=.*Comment","xpath:=//DIV[@id='taskMenuBar']/SPAN[7]").Highlight
		Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=.*Comment","xpath:=//DIV[@id='taskMenuBar']/SPAN[7]").Click
		Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Add CommentSubmit","xpath:=//DIV[@id='claimed_task_form']/DIV[3]/DIV[1]").Highlight
		Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("xpath:=//DIV[normalize-space()='Add Comment']/TEXTAREA[1]").Click
		Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("xpath:=//DIV[normalize-space()='Add Comment']/TEXTAREA[1]").Set DataTable.Value("DiscoReason")
		Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("innertext:=Submit","xpath:=//DIV[@role='button' and normalize-space()='Submit']").Highlight
		Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("innertext:=Submit","xpath:=//DIV[@role='button' and normalize-space()='Submit']").Click
		Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Comments.*","xpath:=//DIV[11]/DIV[1]").Highlight
		
		'Click on Terminate to Terminate the Disconnect Task
		Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=.*Terminate","xpath:=//DIV[@id='taskMenuBar']/SPAN[5]").Highlight
		Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=.*Terminate","xpath:=//DIV[@id='taskMenuBar']/SPAN[5]").Click
		Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Terminate.*","xpath:=//DIV[3]/DIV[@role='dialog'][1]").Highlight
		Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("innertext:=Ok","xpath:=//DIV[@role='button' and normalize-space()='Ok']").Highlight
		Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("innertext:=Ok","xpath:=//DIV[@role='button' and normalize-space()='Ok']").Click
		Else
			Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-captiontext","innertext:=Claimed Task-.*").Exist
			Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=.*Comment","xpath:=//DIV[@id='taskMenuBar']/SPAN[7]").Highlight
			Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=.*Comment","xpath:=//DIV[@id='taskMenuBar']/SPAN[7]").Click
			Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Add CommentSubmit","xpath:=//DIV[@id='claimed_task_form']/DIV[3]/DIV[1]").Highlight
			Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("xpath:=//DIV[normalize-space()='Add Comment']/TEXTAREA[1]").Click
			Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("xpath:=//DIV[normalize-space()='Add Comment']/TEXTAREA[1]").Set DataTable.Value("DiscoReason")
			Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("innertext:=Submit","xpath:=//DIV[@role='button' and normalize-space()='Submit']").Highlight
			Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("innertext:=Submit","xpath:=//DIV[@role='button' and normalize-space()='Submit']").Click
			Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Comments.*","xpath:=//DIV[11]/DIV[1]").Highlight
			
			'Click on Terminate to Terminate the Disconnect Task
			Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=.*Terminate","xpath:=//DIV[@id='taskMenuBar']/SPAN[5]").Highlight
			Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=.*Terminate","xpath:=//DIV[@id='taskMenuBar']/SPAN[5]").Click
			Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Terminate.*","xpath:=//DIV[3]/DIV[@role='dialog'][1]").Highlight
			Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("innertext:=Ok","xpath:=//DIV[@role='button' and normalize-space()='Ok']").Highlight
			Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("innertext:=Ok","xpath:=//DIV[@role='button' and normalize-space()='Ok']").Click
	End If
	
	Else
	If Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innerhtml:=Disconnects Task","xpath:=//TD/DIV[normalize-space()='Disconnects Task']").Exist Then
		'Double Click on Disconnect Task
		Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-table-cell-wrapper","innertext:=Disconnects Task").FireEvent "onmouseover"
		Browser("title:=BPM Console").Page("title:=BPM Console").WebTable("Class Name:=WebTable","class:=v-table-table","text:=Disconnects Task.*").WebElement("Class Name:=WebElement","class:=v-table-cell-wrapper","innertext:=Disconnects Task").Highlight
		Browser("title:=BPM Console").Page("title:=BPM Console").WebTable("Class Name:=WebTable","class:=v-table-table","text:=Disconnects Task.*").WebElement("Class Name:=WebElement","class:=v-table-cell-wrapper","innertext:=Disconnects Task").FireEvent"ondblclick"
		
		If Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-captiontext","innertext:=Candidate Task-.*").Exist Then
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
			
			'Click on Comment & Enter some Text
			Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=.*Comment","xpath:=//DIV[@id='taskMenuBar']/SPAN[7]").Highlight
			Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=.*Comment","xpath:=//DIV[@id='taskMenuBar']/SPAN[7]").Click
			Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Add CommentSubmit","xpath:=//DIV[@id='claimed_task_form']/DIV[3]/DIV[1]").Highlight
			Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("xpath:=//DIV[normalize-space()='Add Comment']/TEXTAREA[1]").Click
			Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("xpath:=//DIV[normalize-space()='Add Comment']/TEXTAREA[1]").Set DataTable.Value("DiscoReason")
			Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("innertext:=Submit","xpath:=//DIV[@role='button' and normalize-space()='Submit']").Highlight
			Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("innertext:=Submit","xpath:=//DIV[@role='button' and normalize-space()='Submit']").Click
			Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Comments.*","xpath:=//DIV[11]/DIV[1]").Highlight
			
			'Click on Terminate to Terminate the Disconnect Task
			Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=.*Terminate","xpath:=//DIV[@id='taskMenuBar']/SPAN[5]").Highlight
			Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=.*Terminate","xpath:=//DIV[@id='taskMenuBar']/SPAN[5]").Click
			Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Terminate.*","xpath:=//DIV[3]/DIV[@role='dialog'][1]").Highlight
			Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("innertext:=Ok","xpath:=//DIV[@role='button' and normalize-space()='Ok']").Highlight
			Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("innertext:=Ok","xpath:=//DIV[@role='button' and normalize-space()='Ok']").Click
			
			Else
				'Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-captiontext","innertext:=Claimed Task-.*").Exist
				Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=.*Comment","xpath:=//DIV[@id='taskMenuBar']/SPAN[7]").Highlight
				Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=.*Comment","xpath:=//DIV[@id='taskMenuBar']/SPAN[7]").Click
				Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Add CommentSubmit","xpath:=//DIV[@id='claimed_task_form']/DIV[3]/DIV[1]").Highlight
				Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("xpath:=//DIV[normalize-space()='Add Comment']/TEXTAREA[1]").Click
				Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("xpath:=//DIV[normalize-space()='Add Comment']/TEXTAREA[1]").Set DataTable.Value("DiscoReason")
				Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("innertext:=Submit","xpath:=//DIV[@role='button' and normalize-space()='Submit']").Highlight
				Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("innertext:=Submit","xpath:=//DIV[@role='button' and normalize-space()='Submit']").Click
				Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Comments.*","xpath:=//DIV[11]/DIV[1]").Highlight
				
				'Click on Terminate to Terminate the Disconnect Task
				Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=.*Terminate","xpath:=//DIV[@id='taskMenuBar']/SPAN[5]").Highlight
				Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=.*Terminate","xpath:=//DIV[@id='taskMenuBar']/SPAN[5]").Click
				Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Terminate.*","xpath:=//DIV[3]/DIV[@role='dialog'][1]").Highlight
				Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("innertext:=Ok","xpath:=//DIV[@role='button' and normalize-space()='Ok']").Highlight
				Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("innertext:=Ok","xpath:=//DIV[@role='button' and normalize-space()='Ok']").Click
		End If
	End If
End If

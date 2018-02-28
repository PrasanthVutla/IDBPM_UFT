'Double Click on Set Priority Task
With Browser("title:=BPM Console").Page("title:=BPM Console")
	.WebTable("Class Name:=WebTable","class:=v-table-table","innertext:=Set Priority.*").WebElement("Class Name:=WebElement","class:=v-table-cell-wrapper","innertext:=Set Priority").Highlight
	If Browser("title:=BPM Console").Page("title:=BPM Console").WebTable("Class Name:=WebTable","class:=v-table-table","innertext:=Set Priority.*").WebElement("Class Name:=WebElement","class:=v-table-cell-wrapper","innertext:=Set Priority").Exist Then
		Reporter.ReportEvent micPass, "Set Priority Task Available", "Task Available"
		.WebTable("Class Name:=WebTable","class:=v-table-table","innertext:=Set Priority.*").WebElement("Class Name:=WebElement","class:=v-table-cell-wrapper","innertext:=Set Priority").FireEvent"ondblclick"
		Else
			Reporter.ReportEvent micFail, "Set Priority Not-Available", "Task Available - Failed"
	End If
End With

'Verify & Claim the Set Priority Task
'Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-captiontext","innertext:=Candidate Task-.*").WaitProperty "innerhtml", "Candidate Task.*"
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-menubar-menuitem","innertext:=.*Claim").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-menubar-menuitem-caption","innertext:=.*Claim").Click
wait 10

Browser("title:=BPM Console").Page("title:=BPM Console").WebTable("Class Name:=WebTable","class:=v-table-table","innertext:=Set Priority.*").WebElement("Class Name:=WebElement","class:=v-table-cell-wrapper","innertext:=Set Priority").FireEvent"ondblclick"

'Selecting the Value from the Dropdown
Browser("title:=BPM Console").Page("title:=BPM Console").WebTable("Class Name:=WebTable","name:=priority","html tag:=TABLE").WebList("Class Name:=WebList","html id:=priority").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebTable("Class Name:=WebTable","name:=priority","html tag:=TABLE").WebList("Class Name:=WebList","html id:=priority").WebButton("Class Name:=WebButton","class:=v-filterselect-button").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebTable("Class Name:=WebTable","name:=priority","html tag:=TABLE").WebList("Class Name:=WebList","html id:=priority").WebButton("Class Name:=WebButton","class:=v-filterselect-button").Click

Browser("title:=BPM Console").Page("title:=BPM Console").WebTable("Class Name:=WebTable","innertext:=123").FireEvent "onmouseover"
Browser("title:=BPM Console").Page("title:=BPM Console").WebTable("Class Name:=WebTable","innertext:=123").WebElement("Class Name:=WebElement","class:=gwt-MenuItem","innertext:=3").Click

'Save SetPriority Task
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-menubar-menuitem","innertext:=.*Save").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-menubar-menuitem","innertext:=.*Save").Click
Wait 10

'Click Complete Task
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-menubar-menuitem","innertext:=.*Complete").FireEvent "onmouseover"
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-menubar-menuitem","innertext:=.*Complete").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","xpath:=//DIV[@id='taskMenuBar']/SPAN[4]","innertext:=.*Complete").CheckProperty "innertext", "Complete"
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","xpath:=//DIV[@id='taskMenuBar']/SPAN[4]","innertext:=.*Complete").Click
Reporter.ReportEvent micDone, "Click on Complete Button - Success", "Clicked on Complete Button"
wait 5

'Click OK on Complete Task
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-window-header","outertext:=Complete Task").Highlight
If Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-window-header","outertext:=Complete Task").Exist Then
	Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("Class Name:=WebButton","innertext:=Ok").Highlight
	Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("Class Name:=WebButton","innertext:=Ok").Click
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("class:=v-nativebutton-caption","innertext:=Disco").FireEvent "onmouseover"
	Reporter.ReportEvent micPass, "Completed Set Priority - Successfully", "Set Priority Completed"
	Else
		Reporter.ReportEvent micFail, "Complete Set Priority - Failed", "Set Priority Not-Completed"
End If



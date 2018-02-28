'Click on Disco Plugin
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("class:=v-nativebutton-caption","innertext:=Disco").Click
If Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-loading-indicator second v-loading-indicator-delay").Exist Then
	wait 20
End If
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("class:=v-nativebutton-caption","innertext:=Disco").FireEvent "onmouseup"
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("class:=v-nativebutton-caption","innertext:=Disco").FireEvent "onmouseover"
Wait 10

'Sync for Disco Plugin
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("class:=v-label v-widget h1 v-label-h1 v-label-undef-w","xpath:=//DIV[normalize-space()='Disco']/DIV[1]").WaitProperty "innertext", "Disco", 1000
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("class:=v-label v-widget h1 v-label-h1 v-label-undef-w","xpath:=//DIV[normalize-space()='Disco']/DIV[1]").CheckProperty "innertext", "Disco", 100
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("class:=v-label v-widget h1 v-label-h1 v-label-undef-w","xpath:=//DIV[normalize-space()='Disco']/DIV[1]").Highlight

'Apply Filter for Work Request Number
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("Class:=v-button v-widget icon-only v-button-icon-only","name:=filterButton").FireEvent"onmouseover"
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("Class:=v-button v-widget icon-only v-button-icon-only","name:=filterButton").Highlight
wait 10
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("Class:=v-button v-widget icon-only v-button-icon-only","name:=filterButton").Click
Wait 5
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("class:=v-nativebutton-caption","innertext:=Disco").FireEvent "onmouseover"
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","html id:=filter-field-select").Click
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","html id:=filter-field-select").WebButton("Class Name:=WebButton","class:=v-filterselect-button").Click
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","innerhtml:=Next").Click
Wait 5
Browser("title:=BPM Console").Page("title:=BPM Console").WebTable("Class Name:=WebTable","innertext:=Submitted ByTask NameWorkRequest Number","name:=WebTable").Highlight
'Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-filterselect-suggestmenu").Highlight 
'Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-filterselect-suggestmenu").FireEvent "onmouseover"
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
If Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-label v-widget label v-label-label v-label-undef-w","innerhtml:=WorkRequest Number equals.*").Exist Then
	Reporter.ReportEvent micPass, "Apply Filter - Sucessful", "Filter Applied"
	'Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-label v-widget label v-label-label v-label-undef-w","innerhtml:=WorkRequest Number equals.*").WaitProperty "innerhtml", "WorkRequest Number equals "&WR_NO
	Else
		Reporter.ReportEvent micFail, "Apply Filter - Failed", "Filter Not-Applied"
End If


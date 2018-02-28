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

'Click & Navigate to Disco Serach Screen
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=.*Search","xpath:=//DIV[@id='discoMenuBar']/SPAN[3]").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=.*Search","xpath:=//DIV[@id='discoMenuBar']/SPAN[3]").Click

'Wait Property for the Search Administration & Blank Search
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("html id:=workRequestLabelId","xpath:=//DIV[@id='workRequestLabelId']").WaitProperty "innertext", "Work Request #", 5000
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("html id:=searchButtonId","xpath:=//DIV[@id='searchButtonId']").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("html id:=searchButtonId","xpath:=//DIV[@id='searchButtonId']").Click
If Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=.*Please enter any search criteria.*","xpath:=//DIV[@role='alert']").Exist Then
	Reporter.ReportEvent micPass, "Blank Search Alert" , "Negative Test - Pass"
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=.*Please enter any search criteria.*","xpath:=//DIV[@role='alert']").Highlight
	Wait 5
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=.*Please enter any search criteria.*","xpath:=//DIV[@role='alert']").FireEvent "onmouseover"
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=.*Please enter any search criteria.*","xpath:=//DIV[@role='alert']").FireEvent "onclick"
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=.*Please enter any search criteria.*","xpath:=//DIV[2]/DIV[@role='alert'][1]").Click
	Else 
		Reporter.ReportEvent micFail, "Search Functionality is Broken", "Negative Test - Failed: Verify the Search Functionality"
			
End If

If Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Search Criteria","xpath:=//DIV[@id='searchCriteriaLabelId']").Exist Then
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Search Criteria","xpath:=//DIV[@id='searchCriteriaLabelId']").Highlight
	Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("html id:=aipSearchId","xpath:=//INPUT[@id='aipSearchId']").Set DataTable.Value("AipId")
	Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("html id:=searchButtonId","xpath:=//DIV[@id='searchButtonId']").Click
	If Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=.*is closed in Vantive and therefore AIPs .*","xpath:=//DIV[@role='alert']").Exist Then
		Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=.*is closed in Vantive and therefore AIPs .*","xpath:=//DIV[@role='alert']").Highlight
		Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=.*is closed in Vantive and therefore AIPs .*","xpath:=//DIV[@role='alert']").FireEvent "onmouseover"
		Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=.*is closed in Vantive and therefore AIPs .*","xpath:=//DIV[@role='alert']").Click
		Reporter.ReportEvent micStatus, "AIP Search Result's with more than 30Day's Closure" , "Search on AIP - PASS"
		Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("html id:=totalMRRLabelId","xpath:=//DIV[@id='totalMRRLabelId']").Highlight
		Else
			Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("html id:=aipSearchId","xpath:=//INPUT[@id='aipSearchId']").Set ""
			Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("html id:=cwpSearchId","xpath:=//INPUT[@id='cwpSearchId']").Set DataTable.Value("CwpId")
			Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("html id:=searchButtonId","xpath:=//DIV[@id='searchButtonId']").Click
			Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=CWP .*","xpath:=//DIV[@role='alert']").Highlight
			Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=CWP .*","xpath:=//DIV[@role='alert']").FireEvent "onmouseover"
			Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=CWP .*","xpath:=//DIV[@role='alert']").Click
			Reporter.ReportEvent micDone, "CWP Search Result's for Disconnect Request" , "Search on CWP's working as Expected - PASS"
	End If
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("html id:=actualETCLabelId","xpath:=//DIV[@id='actualETCLabelId']").Highlight
	Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("html id:=aipSearchId","xpath:=//INPUT[@id='aipSearchId']").Set ""
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=.*Dashboard","xpath:=//DIV[@id='discoMenuBar']/SPAN[1]").Click
End If

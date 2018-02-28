Option Explicit  
Dim StartbillDateError, AipRowId, AipCounter, AipIndex, TaskCwpId, ExcelObj, MySheet
AipIndex = 0 
AipCounter = 1

Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-captiontext","outertext:=AIPS/SERVICE IDS REVIEW").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-captiontext","outertext:=AIPS/SERVICE IDS REVIEW").CheckProperty "outertext", "AIPS/SERVICE IDS REVIEW"
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-captiontext","outertext:=AIPS/SERVICE IDS REVIEW").Click

'If Build AIP is Enabled - Click on the BUILD AIP GRID
If Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("xpath:=//DIV[@id='buildAipGridButtonMenuBar']/SPAN","innertext:=Build Aip Grid","visible:=True").Exist Then
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("xpath:=//DIV[@id='buildAipGridButtonMenuBar']/SPAN","innertext:=Build Aip Grid","visible:=True").Highlight
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("xpath:=//DIV[@id='buildAipGridButtonMenuBar']/SPAN","innertext:=Build Aip Grid","visible:=True").CheckProperty "innertext", "Build Aip Grid"
	'Build AIP -  Enabled successful - Click on the BUILD AIP
	Reporter.ReportEvent micPass, "BuildAIP", "Build AIP is Enabled"   
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("xpath:=//DIV[@id='buildAipGridButtonMenuBar']/SPAN","innertext:=Build Aip Grid","visible:=True").Click
	wait 5
	
	'AIP's Start Bill Date Not Available - Handle the Error mesg
	If Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("xpath:=//DIV[2]/DIV[@role='alert'][1]/DIV[1]/DIV[1]").Exist Then
		Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("xpath:=//DIV[2]/DIV[@role='alert'][1]/DIV[1]/DIV[1]").Highlight
		Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("xpath:=//DIV[2]/DIV[@role='alert'][1]/DIV[1]/DIV[1]").FireEvent "onmouseover"
		Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("xpath:=//DIV[2]/DIV[@role='alert'][1]/DIV[1]/DIV[1]").CaptureBitmap "StartbillDate.png", True
		Set StartbillDateError = Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("xpath:=//DIV[2]/DIV[@role='alert'][1]/DIV[1]/DIV[1]")
		'msgbox StartbillDateError.GetROProperty("innertext")
		Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("xpath:=//DIV[2]/DIV[@role='alert'][1]/DIV[1]/DIV[1]").Click
		Reporter.ReportEvent micFail, "This test aborted at line 20/Build AIP - Fialure:","AIP's Start Bill Date missing - Pls Work With DBA","StartbillDate.png"
		ExitTest
		
		Else
		'AIP Grid
		Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("xpath:=//DIV[3]/DIV[@role='dialog'][1]/DIV[1]/DIV[1]/DIV[2]").Highlight
		Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("htmlid:=Complete AIP Grid","xpath:=//DIV[@id='Complete AIP Grid']").Highlight
		Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("htmlid:=Complete AIP Grid","xpath:=//DIV[@id='Complete AIP Grid']").Click
		wait 3

		'Handle Complete AIP Grid Alert with out filling the Data - Failure Exception
		Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","xpath:=//DIV[2]/DIV[@role='alert']").Highlight
		Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","xpath:=//DIV[2]/DIV[@role='alert']").FireEvent "onmouseover"
		Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","xpath:=//DIV[2]/DIV[@role='alert']").Click
		wait 2
	End If

	'Select AD Justification Code, Contract Status, ETC Eligible - Implement For Loop
	Browser("title:=BPM Console").Page("title:=BPM Console").WebTable("xpath:=//DIV[@id='aipGridTable']/DIV[2]/DIV[1]/TABLE[1]").Highlight
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("xpath:=//DIV[@role='dialog']//DIV[1]/DIV[5]/DIV[1]/DIV[1]/DIV[1]/DIV[1]").Highlight
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("xpath:=//DIV[@role='dialog']//DIV[1]/DIV[5]/DIV[1]/DIV[1]/DIV[1]/DIV[1]").CheckProperty "innertext", "Data*CopyTo*Row From*Row To*Copy"
	Set AipRowId = Browser("title:=BPM Console").Page("title:=BPM Console").WebTable("xpath:=//DIV[@id='aipGridTable']/DIV[2]/DIV[1]/TABLE[1]")
	'msgbox AipRowId.GetROProperty("Rows")
	
	'Validate the Copy Function
	Browser("title:=BPM Console").Page("title:=BPM Console").WebList("Class Name:=WebList","html id:=financeRollUp","index:="&AipIndex).Highlight
	Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("html id:=DataId","xpath:=//INPUT[@id='DataId']").Highlight
	Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("html id:=DataId","xpath:=//INPUT[@id='DataId']").Set DataTable.Value("financeRollUp")
	Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("xpath:=//DIV[@id='copyToId']/DIV[@role='button']").Click
	Browser("title:=BPM Console").Page("title:=BPM Console").WebTable("Class Name:=WebTable","innertext:=financeRollUp.*").Highlight
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","innerhtml:=financeRollUp").Click
	Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("xpath:=//DIV[@id='rowFromId']/DIV[@role='button']").Click
	Browser("title:=BPM Console").Page("title:=BPM Console").WebTable("Class Name:=WebTable","innertext:=1.*").Highlight
	Browser("title:=BPM Console").Page("title:=BPM Console").WebTable("Class Name:=WebTable","innertext:=1.*").FireEvent "onmouseover"
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("xpath:=//TR/TD[@role='listitem' and normalize-space()='1']","innertext:=1").Click
	Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("xpath:=//DIV[@id='rowToId']/DIV[@role='button']").Click
	Browser("title:=BPM Console").Page("title:=BPM Console").WebTable("Class Name:=WebTable","innertext:=1.*").Highlight
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("xpath:=//TR/TD[@role='listitem' and normalize-space()='1']","innertext:=1").Click
	Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("html id:=copyId","xpath:=//DIV[@id='copyId']").Highlight
	Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("html id:=copyId","xpath:=//DIV[@id='copyId']").CheckProperty "html id", "copyId"
	Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("html id:=copyId","xpath:=//DIV[@id='copyId']").Click
	wait 5
	Browser("title:=BPM Console").Page("title:=BPM Console").WebList("Class Name:=WebList","html id:=financeRollUp","index:="&AipIndex).Highlight
	
	'Check the condition at the beginning of the loop
	Do Until AipIndex = AipRowId.GetROProperty("Rows")
		'msgbox "Result of Aip Index at Begining: " &AipIndex
		'Statments to Execute under Loop
		Browser("title:=BPM Console").Page("title:=BPM Console").WebList("Class Name:=WebList","html id:=adJustificationCode","index:="&AipIndex).Highlight
		Browser("title:=BPM Console").Page("title:=BPM Console").WebList("Class Name:=WebList","html id:=adJustificationCode","index:="&AipIndex).WebButton("xpath:=//DIV[@id='adJustificationCode']/DIV[@role='button']").Highlight
		Browser("title:=BPM Console").Page("title:=BPM Console").WebList("Class Name:=WebList","html id:=adJustificationCode","index:="&AipIndex).WebButton("xpath:=//DIV[@id='adJustificationCode']/DIV[@role='button']").Click
		Browser("title:=BPM Console").Page("title:=BPM Console").WebTable("Class Name:=WebTable","innertext:=Backlog.*","index:="&AipIndex).Highlight
		Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","innerhtml:=Backlog","index:="&AipIndex).Click
		wait 20
		
		Browser("title:=BPM Console").Page("title:=BPM Console").WebList("Class Name:=WebList","html id:=etcEligible","index:="&AipIndex).Highlight
		Browser("title:=BPM Console").Page("title:=BPM Console").WebList("Class Name:=WebList","html id:=etcEligible","index:="&AipIndex).WebButton("xpath:=//DIV[@id='etcEligible']/DIV[@role='button']").Highlight
		Browser("title:=BPM Console").Page("title:=BPM Console").WebList("Class Name:=WebList","html id:=etcEligible","index:="&AipIndex).WebButton("xpath:=//DIV[@id='etcEligible']/DIV[@role='button']").Click
		Browser("title:=BPM Console").Page("title:=BPM Console").WebTable("Class Name:=WebTable","innertext:=No.*","index:="&AipIndex).Highlight
		Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","innerhtml:=No","index:="&AipIndex).Click
		wait 20
		
		Browser("title:=BPM Console").Page("title:=BPM Console").WebList("Class Name:=WebList","html id:=contractStatus","index:="&AipIndex).Highlight
		Browser("title:=BPM Console").Page("title:=BPM Console").WebList("Class Name:=WebList","html id:=contractStatus","index:="&AipIndex).WebButton("xpath:=//DIV[@id='contractStatus']/DIV[@role='button']").Highlight
		Browser("title:=BPM Console").Page("title:=BPM Console").WebList("Class Name:=WebList","html id:=contractStatus","index:="&AipIndex).WebButton("xpath:=//DIV[@id='contractStatus']/DIV[@role='button']").Click
		Browser("title:=BPM Console").Page("title:=BPM Console").WebTable("Class Name:=WebTable","innertext:=At end of Original.*","index:="&AipIndex).Highlight
		Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","innerhtml:=Next").Click
		Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","innerhtml:=NCF","index:="&AipIndex).Click
		wait 10
		Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("htmlid:=saveChangesButton","xpath:=//DIV[@id='saveChangesButton']").Highlight
		Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("htmlid:=saveChangesButton","xpath:=//DIV[@id='saveChangesButton']").Click
		wait 15
		 
		AipIndex = AipIndex + AipCounter 'Increment the counter
		'msgbox "Result in loop before Aip Incerment: " &AipIndex
		AipCounter  = AipCounter + 0
		'Display the result
		'msgbox "Result in Loop for AIP Index: " & AipIndex
			If AipIndex = AipRowId.GetROProperty("Rows") Then
    			Exit Do 'Exit inner loop
    		End If
	Loop
	
	'Save the AIP Records
	Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("htmlid:=Complete AIP Grid","xpath:=//DIV[@id='Complete AIP Grid']").Highlight
	Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("htmlid:=Complete AIP Grid","xpath:=//DIV[@id='Complete AIP Grid']").Click
	wait 8
	If Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Info: .*","xpath:=//DIV[@role='alert']").Exist Then
		Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Info: .*","xpath:=//DIV[@role='alert']").Highlight
		Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Info: .*","xpath:=//DIV[@role='alert']").FireEvent "onmouseover"
		Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Info: .*","xpath:=//DIV[@role='alert']").Click
		Reporter.ReportEvent micDone, "BUILD AIP Task Completed","BUILD AIP Task Completed"
	End If
	Reporter.ReportEvent micPass, "BUILD AIP", "Build AIP Grid is Completed Successfull - Verify for the View AIP" 
	wait 20
	
	'Verify the View AIP Grid Button after completing the BUILD AIP
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("class:=v-nativebutton-caption","innertext:=Disco").FireEvent "onmouseup"
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-captiontext","outertext:=CWP ATTRIBUTES").Highlight
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-captiontext","outertext:=CWP ATTRIBUTES").Click
	wait 3
	'####Handled as part of Claim&Disconnect Task Action##### So Commenting - Writting the CWP ID to Xls File
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-captiontext","outertext:=AIPS/SERVICE IDS REVIEW").Highlight
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-captiontext","outertext:=AIPS/SERVICE IDS REVIEW").Click
	'Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","innertext:=View Aip Grid").FireEvent "onmouseover"
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("xpath:=//DIV[@id='buildAipGridButtonMenuBar']/SPAN","innertext:=View Aip Grid").Highlight
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("xpath:=//DIV[@id='buildAipGridButtonMenuBar']/SPAN","innertext:=View Aip Grid","visible:=True").CheckProperty "innertext", "View Aip Grid"
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("xpath:=//DIV[@id='buildAipGridButtonMenuBar']/SPAN","innertext:=View Aip Grid","visible:=True").Click
	Wait 15
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("xpath:=//DIV[3]/DIV[@role='dialog'][1]/DIV[1]/DIV[1]/DIV[2]").Highlight
	Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("xpath:=//DIV[3]/DIV[@role='dialog'][1]/DIV[1]/DIV[1]/DIV[@role='button'][2]").Click
	Wait 5

'Build AIP -  Completed, Moved View AIP Grid
ElseIf Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("xpath:=//DIV[@id='buildAipGridButtonMenuBar']/SPAN","innertext:=View Aip Grid","visible:=True").Exist Then
	Reporter.ReportEvent micPass, "ViewAIP", "View AIP Flow Enabled" 
	
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-captiontext","outertext:=CWP ATTRIBUTES").Highlight
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-captiontext","outertext:=CWP ATTRIBUTES").Click
	wait 5
	
	'####Handled as part of Claim&Disconnect Task Action##### So Commenting - Writting the CWP ID to Xls File
	'Click on AIPS/SERVICE IDS REVIEW Tab
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-captiontext","outertext:=AIPS/SERVICE IDS REVIEW").Highlight
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-captiontext","outertext:=AIPS/SERVICE IDS REVIEW").CheckProperty "outertext", "AIPS/SERVICE IDS REVIEW"
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-captiontext","outertext:=AIPS/SERVICE IDS REVIEW").Click
	
	'Open View AIP Grid and Close it
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("xpath:=//DIV[@id='buildAipGridButtonMenuBar']/SPAN","innertext:=View Aip Grid","visible:=True").Highlight
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("xpath:=//DIV[@id='buildAipGridButtonMenuBar']/SPAN","innertext:=View Aip Grid","visible:=True").Click
	wait 10
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("xpath:=//DIV[3]/DIV[@role='dialog'][1]/DIV[1]/DIV[1]/DIV[2]").Highlight
	Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("xpath:=//DIV[3]/DIV[@role='dialog'][1]/DIV[1]/DIV[1]/DIV[@role='button'][2]").Click
	wait 5
		
Else
	'Build AIP Grid is Not Enabled
	Reporter.ReportEvent micFail, "BuildAIP", "Build AIP is Not Enabled"
	
End If
'ExcelObj.Application.Quit
'Set ExcelObj=nothing

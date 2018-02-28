'Click on Comment & Enter some Text
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=.*Comment","xpath:=//DIV[@id='taskMenuBar']/SPAN[7]").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=.*Comment","xpath:=//DIV[@id='taskMenuBar']/SPAN[7]").Click
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Add CommentSubmit","xpath:=//DIV[@id='claimed_task_form']/DIV[3]/DIV[1]").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("xpath:=//DIV[normalize-space()='Add Comment']/TEXTAREA[1]").Click
Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("xpath:=//DIV[normalize-space()='Add Comment']/TEXTAREA[1]").Set DataTable.Value("DiscoReason")
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("innertext:=Submit","xpath:=//DIV[@role='button' and normalize-space()='Submit']").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("innertext:=Submit","xpath:=//DIV[@role='button' and normalize-space()='Submit']").Click
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Comments.*","xpath:=//DIV[11]/DIV[1]").Highlight

'Verify the Send Email & Send it
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("xpath:=//DIV[@id='taskMenuBar']/SPAN[8]").Highlight
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("xpath:=//DIV[@id='taskMenuBar']/SPAN[8]").Click
If Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("xpath:=//DIV[2]/DIV[@role='alert'][1]","innertext:=Error: Please fill in all the re.*").Exist Then
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("xpath:=//DIV[2]/DIV[@role='alert'][1]","innertext:=Error: Please fill in all the re.*").Highlight
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("xpath:=//DIV[2]/DIV[@role='alert'][1]","innertext:=Error: Please fill in all the re.*").FireEvent "onmouseover"
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("xpath:=//DIV[2]/DIV[@role='alert'][1]","innertext:=Error: Please fill in all the re.*").Click
	Reporter.ReportEvent micFail, "SEND EMAIL", "SEND EMAIL - Build/View AIP is Completed, Issue with Code"
	Else
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Email Editor.*","xpath:=//DIV[3]/DIV[@role='dialog'][1]/DIV[1]/DIV[1]").Highlight
	'Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Email Editor.*","xpath:=//DIV[3]/DIV[@role='dialog'][1]/DIV[1]/DIV[1]").CheckProperty "innertext", "Email Editor.*"
	Browser("title:=BPM Console").Page("title:=BPM Console").WebList("html id:=templateComboboxId","xpath:=//DIV[@id='templateComboboxId']").WebButton("xpath:=//DIV[@id='templateComboboxId']/DIV[@role='button']").Highlight
	Browser("title:=BPM Console").Page("title:=BPM Console").WebList("html id:=templateComboboxId","xpath:=//DIV[@id='templateComboboxId']").WebButton("xpath:=//DIV[@id='templateComboboxId']/DIV[@role='button']").Click
	Browser("title:=BPM Console").Page("title:=BPM Console").WebTable("Class Name:=WebTable","innertext:=US Temp.*").Highlight
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","innerhtml:=Emea template").Click
	wait 5
	Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("xpath:=//INPUT[@id='toId']").Set DataTable.Value("TechContEmail")
	Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("htmlid:=sendEmailButtonId","xpath:=//DIV[@id='sendEmailButtonId']").Highlight
	'Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("htmlid:=sendEmailButtonId","xpath:=//DIV[@id='sendEmailButtonId']").CheckProperty "xpath", "//DIV[@id='sendEmailButtonId']"
	Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("htmlid:=sendEmailButtonId","xpath:=//DIV[@id='sendEmailButtonId']").Click
	wait 8
	If Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Info: .*","xpath:=//DIV[@role='alert']").Exist Then
		Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Info: .*","xpath:=//DIV[@role='alert']").Highlight
		Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Info: .*","xpath:=//DIV[@role='alert']").FireEvent "onmouseover"
		Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Info: .*","xpath:=//DIV[@role='alert']").Click
		Reporter.ReportEvent micPass, "Email Sent Successful", "Email Sent Successful"
		ElseIf Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Error: Error encountered while sending the email.*","xpath:=//DIV[@role='alert']").Exist Then
				Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Error: Error encountered while sending the email.*","xpath:=//DIV[@role='alert']").Highlight
				Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Error: Error encountered while sending the email.*","xpath:=//DIV[@role='alert']").Click
				Reporter.ReportEvent micWarning, "Send Email Error: Check WebSter/Email Service is Down" , "Set the Port Number to 5000 & Resent the Email"
				Wait 4
				'Handle Webster to Set the Port Value to 5000
				Browser("CreationTime:=0").OpenNewTab()
				Browser("CreationTime:=1").Sync
 				'Load some web page in the new tab
				Browser("CreationTime:=1").Navigate "http://dl1tomappdv03b.it.savvis.net:51201/webster/Messages"
				Browser("title:=.*/webster/Messages.*").Page("url:=.*/webster/Messages.*").Link("innertext:=Home","xpath:=//H3[1]/A[1]").Highlight
				Browser("title:=.*/webster/Messages.*").Page("url:=.*/webster/Messages.*").Link("innertext:=Home","xpath:=//H3[1]/A[1]").Click
				Browser("title:=.*/webster/Home.*").Page("url:=.*/webster/Home.*").WebEdit("xpath:=//INPUT[@id='serverPort']").Highlight
				Browser("title:=.*/webster/Home.*").Page("url:=.*/webster/Home.*").WebEdit("xpath:=//INPUT[@id='serverPort']").Set "5000"
				Browser("title:=.*/webster/Home.*").Page("url:=.*/webster/Home.*").WebButton("name:=START","type:=submit","xpath:=//TD[3]/INPUT[1]").Highlight
				Browser("title:=.*/webster/Home.*").Page("url:=.*/webster/Home.*").WebButton("name:=START","type:=submit","xpath:=//TD[3]/INPUT[1]").Click
				Wait 3
					'Browser("title:=.*/webster/Home.*").Close
								
				'Handled Close the Send email and then Open again
				Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("xpath:=//DIV[@id='taskMenuBar']/SPAN[8]").Highlight
				Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("xpath:=//DIV[@id='taskMenuBar']/SPAN[8]").Click
				Browser("title:=BPM Console").Page("title:=BPM Console").WebList("html id:=templateComboboxId","xpath:=//DIV[@id='templateComboboxId']").WebButton("xpath:=//DIV[@id='templateComboboxId']/DIV[@role='button']").Click
				Browser("title:=BPM Console").Page("title:=BPM Console").WebTable("Class Name:=WebTable","innertext:=US Temp.*").Highlight
				Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","innerhtml:=Emea template").Click
				Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("xpath:=//INPUT[@id='toId']").Set DataTable.Value("TechContEmail")
				Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("htmlid:=sendEmailButtonId","xpath:=//DIV[@id='sendEmailButtonId']").Highlight
				Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("htmlid:=sendEmailButtonId","xpath:=//DIV[@id='sendEmailButtonId']").Click
				wait 3
					If Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Info: .*","xpath:=//DIV[@role='alert']").Exist Then
						Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Info: .*","xpath:=//DIV[@role='alert']").Highlight
						Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Info: .*","xpath:=//DIV[@role='alert']").FireEvent "onmouseover"
						Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Info: .*","xpath:=//DIV[@role='alert']").Click
						Reporter.ReportEvent micPass, "Send Email Successful - After Webster Port Set", "Email Sent Successful"
					End If
				Else
					Reporter.ReportEvent micFail, "Send Email Error: Check Email Service" , "Ending the Test - Check with Dev"
					Browser("title:=BPM Console").CloseAllTabs
					ExitTest
	End If
	
	'Complete the Disconnect Task - Click on the Disconnect Button
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("xpath:=//DIV[@id='taskMenuBar']/SPAN[4]").FireEvent "onmouseover"
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("xpath:=//DIV[@id='taskMenuBar']/SPAN[4]").Highlight
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("xpath:=//DIV[@id='taskMenuBar']/SPAN[4]").CheckProperty "innertext", "Complete"
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("xpath:=//DIV[@id='taskMenuBar']/SPAN[4]").Click
	Wait 5
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-window-header","outertext:=Complete Task").Highlight
	Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("Class Name:=WebButton","innertext:=Ok").Highlight
	Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("Class Name:=WebButton","innertext:=Ok").Click
	If Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Info: .*","xpath:=//DIV[@role='alert']").Exist Then
		Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Info: .*","xpath:=//DIV[@role='alert']").Highlight
		Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Info: .*","xpath:=//DIV[@role='alert']").FireEvent "onmouseover"
		Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("innertext:=Info: .*","xpath:=//DIV[@role='alert']").Click
		Reporter.ReportEvent micPass, "Disconnect Task", "Disconnect Task Completed Successful"
	End If
	wait 4
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("class:=v-nativebutton-caption","innertext:=Disco").FireEvent "onmouseover"
End If


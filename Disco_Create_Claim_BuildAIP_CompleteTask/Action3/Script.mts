'Login to BPM
Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("html id:=login_username","class:=v-textfield v-widget.*").Set ""
Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("html id:=login_username","class:=v-textfield v-widget.*").Set DataTable.Value("UserId")
Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("class:=v-textfield v-widget","html id:=login_password").Set ""
Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("class:=v-textfield v-widget","html id:=login_password").SetSecure DataTable.Value("Password")
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("class:=v-button v-widget default v-button-default","html id:=login_signin_button").Click
If Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-loading-indicator second v-loading-indicator-delay").Exist Then
	wait 20
End If
Setting("DefaultTimeout") = 10000

'Sync for Disco Plugin
If Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-label v-widget h1 v-label-h1 v-label-undef-w").Exist Then
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-label v-widget h1 v-label-h1 v-label-undef-w").CheckProperty "innertext", "Disco", 100
	wait 25
	Reporter.ReportEvent micPass, "Login to BPM: Successful", "Logged in to BPM Console"
	Else
		Reporter.ReportEvent micFail, "Login to BPM: Failed", "Verify the Credential/Application"
End If




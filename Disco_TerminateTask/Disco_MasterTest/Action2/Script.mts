'Login to BPM
Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("html id:=login_username","class:=v-textfield v-widget.*").Set ""
Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("html id:=login_username","class:=v-textfield v-widget.*").Set DataTable.Value("UserId")
Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("class:=v-textfield v-widget","html id:=login_password").Set ""
Browser("title:=BPM Console").Page("title:=BPM Console").WebEdit("class:=v-textfield v-widget","html id:=login_password").SetSecure DataTable.Value("Password")
Browser("title:=BPM Console").Page("title:=BPM Console").WebButton("class:=v-button v-widget default v-button-default","html id:=login_signin_button").Click
wait 20
Setting("DefaultTimeout") = 60000

'Click Disco Plugin
'Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("class:=v-nativebutton-caption","innertext:=Disco").Click
If Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-loading-indicator second v-loading-indicator-delay").Exist Then
	Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-loading-indicator second v-loading-indicator-delay").Highlight
	wait 15
End If

'Sync for Disco Plugin
Browser("title:=BPM Console").Page("title:=BPM Console").WebElement("Class Name:=WebElement","class:=v-label v-widget h1 v-label-h1 v-label-undef-w").WaitProperty "innertext", "Disco", 100

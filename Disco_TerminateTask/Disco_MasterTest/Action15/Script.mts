Dim oOutlook, oEmail, vEmailTo, vEmailCC

vEmailTo = ""    'replace  this with your email id
vEmailCC = "pvutla@centurylink.com"
vEmailSubject = "UFT Test Results - BAU Activity"

Set oOutlook = CreateObject("Outlook.Application")
Set oEmail = oOutlook.CreateItem(0)

oEmail.To = vEmailTo
oEmail.CC = vEmailCC
oEmail.Subject = vEmailSubject

oEmail.HTMLBody = "<HTML>"&_
"<Body>"&_
"<p Style=""background-color:red""><b>BUA Test Results using UFT Automation</b></p>"&_
"<p><strong><Font color=""Orange"" size = 16>Test Results for Disco</Font></strong></p>"&_
"<p><em>Scenario's Outlined for Execution & It's Status as Below</em></p>"&_
"<p><Font color=""green"" size = 14<i>Test PASS</Font></i></p>"&_
"<p><Font color=""red"" size = 14><i>Test FAIL</Font></i></p>"&_
"<p>This is Automated Test Results Email</p>"&_
"</Body>"&_
"</HTML>"

wait 2
oEmail.Send
wait 2

Set oEmail = Nothing
Set oOutlook = Nothing

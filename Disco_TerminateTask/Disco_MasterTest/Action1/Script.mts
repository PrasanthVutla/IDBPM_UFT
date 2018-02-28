'Open Browser & Navigate to Test URL
Dim URL
URL = "http://bpm-qa.it.savvis.net/bpm-console-web/"
SystemUtil.Run "Chrome",URL

'Variables
Dim curDateTime, curDate, curTime
curDateTime = Now
curDate = Date
curTime = Time

'Handle Disable Developer Extensions
Browser("BPM Console").InsightObject("InsightObject").Click @@ hightlight id_;_2_;_script infofile_;_ZIP::ssf24.xml_;_


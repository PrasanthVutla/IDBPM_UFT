'Open Browser & Navigate to Test URL
Datatable.Import "C:\IDBPM_UFT\Test_Data\Test_Data.xls"
SystemUtil.Run "Chrome",DataTable.Value ("URL")

'Variables
Dim curDateTime, curDate, curTime
curDateTime = Now
curDate = Date
curTime = Time

'Handle Disable Developer Extensions
Browser("BPM Console").InsightObject("InsightObject").Click

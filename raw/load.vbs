On Error Resume Next

Function pad(num)
	if num < 10 then
		num="0"&num
	end if 
	pad=num
End Function

Function formatDate(dt)
	ss=pad(Second(dt))
	mm=pad(Minute(dt))
	hh=pad(hour(dt))
	dd=pad(Day(dt))
	mth=pad(Month(dt))
	YY=Year(dt)
	formatDate=dd&"/"&mth&"/"&YY&" "&hh&":"&mm&":"&ss
End Function

strComputer = "."
query = "Select * from Win32_PerfFormattedData_PerfOS_System "
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery(query,,48)
logMessage = ""
sep = "," 
For Each objItem in colItems
	REM logMessage = FormatDateTime(Now(),2) & " " & FormatDateTime(Now(),4)  & sep 
	logMessage = formateDate(Now()) & sep 
	logMessage = logMessage &  objItem.Name & sep 
	logMessage = logMessage &  objItem.ProcessorQueueLength & sep 
	logMessage = logMessage &  objItem.Threads 
	logMessage = logMessage &  objItem.SystemCallsPerSec & sep 
	WScript.Echo logMessage
	
Next

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
	formatDate=YY&"/"&mth&"/"&dd&" "&hh&":"&mm&":"&ss
End Function

strComputer = "."
sep = ","
Set wshNetwork = CreateObject( "WScript.Network" )
host = WshNetwork.ComputerName
timestamp = formatDate(Now())


Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" _
& strComputer & "\root\cimv2")

Set colService = objWMIService.ExecQuery("SELECT DisplayName,Status,State,ProcessId,ServiceType,StartName,Started,StartMode,Name,ExitCode FROM Win32_Service")

For Each objService in colService
   WSCript.Echo timestamp _
   & sep & host _
   & sep & objService.DisplayName _
   & sep & objService.Status _
   & sep & objService.State _
   & sep & objService.ProcessId _
   & sep & objService.ServiceType _
   & sep & objService.StartName _
   & sep & objService.Started _
   & sep & objService.StartMode _
   & sep & objService.Name _
   & sep & objService.ExitCode
Next

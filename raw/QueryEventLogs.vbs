'----------------------------------------------------------------------------------------------------------------------------
'Script Name : QueryEventLogs.vbs
'Author      : Matthew Beattie
'Created     : 16/09/09
'Description : This script queries the event log for...whatever you want it to! Just set the event log name and event ID's!
'
'ak-- 28/11/13 -- heavily modified for use in Logscape.
' Now it takes two optional parameters: LogName and fromDateTime, like this:
' cscript.exe QueryEventLogs.vbs Application '26/11/2013 10:19:10'
' - if not supplied, then the default values are used (System and Now()-1min)
'
'----------------------------------------------------------------------------------------------------------------------------
'Initialization  Section
'----------------------------------------------------------------------------------------------------------------------------
Option Explicit
Const ForReading   = 1
Const ForWriting   = 2
Const ForAppending = 8
Const localeGB = "en-gb"
Dim objDictionary, objFSO, wshShell, wshNetwork
Dim ipAddress, macAddress, item, messageType, message
Dim args, logName, OSVersionMajor, currentTimeZone
Dim filter
Dim dtDay, dtMonth, dtYear, dtTime, dtSec
Dim localeHere
Dim fromDateTime, UntilDateTime, strFromDateTime, strUntilDateTime

On Error Resume Next
	Set objDictionary = NewDictionary
	Set objFSO        = CreateObject("Scripting.FileSystemObject")
	Set wshShell      = CreateObject("Wscript.Shell")
	Set wshNetwork    = CreateObject("Wscript.Network")
	
	Dim properties
	Set properties = getPropertiesFromArguments()
	logName=getProperty(properties,"eventLog","System")
	filter=getProperty(properties,"excludeEventCodes","")
		
	If Err.Number <> 0 Then
		Wscript.Quit
	End If
	
	'Parsing time arguments and slicing - only request the events during the whole interval, 
	' from the start of the minute before current to the beginning of the current minute
	localeHere = GetLocale()
	strFromDateTime=getProperty(properties,"startDate","")
	strEndDateTime=getProperty(properties,"endDate","")
	If strFromDateTime="" then 
		'not given From Time - then default it to -1 min and also 
		'define the To time, as current - both up to 1 min
		
		fromDateTime = DateAdd("n", -1, Now())
		dtSec = Second(fromDateTime)
		fromDateTime = DateAdd("s", -dtSec, fromDateTime)
		UntilDateTime = DateAdd("n", 1, fromDateTime)

		' Debuggig lines 
		'fromDateTime = DateAdd("s", -3600*24*24, fromDateTime)
		'UntilDateTime = DateAdd("n", 600*24*7, fromDateTime)		
		
	Else 'specific time(s) given on the command line
		strFromDateTime =strFromDateTime
		'Need to convert the DateTime here from the expected format (in en-gb locale: dd/MM/YYYY HH:mm:ss) 
		' into whatever is local to the machine
		SetLocale(localeGB)
		fromDateTime = CDate(strFromDateTime)
		SetLocale(localeHere)
		'WScript.Echo "From time given was: " & args.Item(1) & " and interpreted locally as: " & fromDateTime
		If strEndDateTime="" then 
			'From time was given but not the To time - assume now.
			UntilDateTime = CDate(Now())
		Else
			strUntilDateTime = strEndDateTime
			'also convert to local time
			SetLocale(localeGB)
			UntilDateTime = CDate(strUntilDateTime)
			SetLocale(localeHere)
			'WScript.Echo "To time given was: " & args.Item(2) & " and interpreted locally as: " & untilDateTime
		End If
	End If
	
	'WScript.Echo "Requesting data from: " & fromDateTime & " to: " & untilDateTime
On Error Goto 0
'----------------------------------------------------------------------------------------------------------------------------
'Main Processing Section
'----------------------------------------------------------------------------------------------------------------------------
On Error Resume Next
	QueryOSdetails OSVersionMajor, currentTimeZone
	ProcessScript OSVersionMajor
	If Err.Number <> 0 Then
		MsgBox BuildError("Processing Script"), vbCritical, scriptBaseName
		Wscript.Quit
	End If
On Error Goto 0


Function getProperty(properties,key,defaultValue)
	Dim value
	If properties.Exists(key) = true Then
		value=properties(key)
	else
		value=defaultValue
	End If
	getProperty=value
End Function

'----------------------------------------------------------------------------------------------------------------------------
'Functions Processing Section
'----------------------------------------------------------------------------------------------------------------------------
'Name       : getPropertiesFromArguments -> Creates a Dictionary from logscape properties
'Parameters : None          ->
'Return     : Dictionary    ->
'----------------------------------------------------------------------------------------------------------------------------
Function getPropertiesFromArguments()
	Dim i
	Dim properties
	Set properties=CreateObject("Scripting.Dictionary")
	For i=0 to WScript.Arguments.Count() - 1 

		If InStr(WScript.Arguments(i),"=") Then
			Dim elems
			elems=Split(WScript.Arguments(i),"=")
			properties.Add elems(0),elems(1)
		End If 
	Next 
	Set getPropertiesFromArguments=properties 
End Function


'----------------------------------------------------------------------------------------------------------------------------
'Functions Processing Section
'----------------------------------------------------------------------------------------------------------------------------
'Name       : ProcessScript -> Primary Function that controls all other script processing.
'Parameters : None          ->
'Return     : None          ->
'----------------------------------------------------------------------------------------------------------------------------
Function ProcessScript(OS)
	Dim events, hostName, i
	'hostName      = wshNetwork.ComputerName
	hostName      = "."						'ak-- only use locally
	'-------------------------------------------------------------------------------------------------------------------------
	'Construct part of the WMI Query to account for searching multiple eventID's
	'-------------------------------------------------------------------------------------------------------------------------
	If Not QueryEventLog(events, hostName, logName,filter) Then
		Exit Function
	End If
	'-------------------------------------------------------------------------------------------------------------------------
	'Log the scripts results to the scripts
	'-------------------------------------------------------------------------------------------------------------------------
	select Case OS
	Case 5 'Windows version major 5 is for Windows 2003 and XP
		' Printing out event records in normal order
		For i = 1 To events.Count Step 1
			LogMessage events.Item(CStr(i))
		Next
	Case 6 'for now it is Windows version major 6 - Windows 7 and 2008. 
		' Printing out event records in reverse order, as WQL query with conditions does return the results in reverse order
		' and does not allow for ordering... :-(
		For i = events.Count To 1 Step -1
			LogMessage events.Item(CStr(i))
		Next
	Case Else
		'do nothing - don't know how to process yet
		Exit Function
	End Select
End Function


Function filterExcludeEvents(excludes,code)

	
	Dim res
	res=0
	If InStr(excludes,","&code) > 0 Then
		res=1
	End If
	
	If InStr(excludes,code&",") > 0 Then
		res=1
	End If
	
	If excludes=CStr(code) Then 
		res=1
		
	End If 
	
	filterExcludeEvents=res
End Function
'----------------------------------------------------------------------------------------------------------------------------
'Name       : QueryEventLog -> Primary Function that controls all other script processing.
'Parameters : results       -> Input/Output : Variable assigned to an array of results from querying the event log.
'           : hostName      -> String containing the hostName of the system to query the event log on.
'           : logName       -> String containing the name of the Event Log to query on the system.
'           : eventNumbers  -> Array containing the EventID's (eventCode) to search for within the event log.
'           : fromDateTime -> Date\Time containing the date to finish searching at.
'           : minutes       -> Integer containing the number of minutes to subtract from the startDate to begin the search.
'Return     : QueryEventLog -> Returns True if the event log was successfully queried otherwise returns False.
'----------------------------------------------------------------------------------------------------------------------------
Function QueryEventLog(eventsDict, hostName, logName ,excludeCodes)
   Dim wmi, query, result, results, eventInfo, cShortName, cFullName
   Dim wmiDateTime, strFrom, strUntil, DateTime24, errorCount, i
   QueryEventLog = False
   errorCount    = 0
   '-------------------------------------------------------------------------------------------------------------------------
   'Construct part of the WMI Query to account for searching multiple eventID's
   '-------------------------------------------------------------------------------------------------------------------------
   query = "Select * from Win32_NTLogEvent Where Logfile = " & SQ(logName) 
   'WScript.Echo "query so far: " & query 
   On Error Resume Next
      Set eventsDict = NewDictionary
      If Err.Number <> 0 Then
         LogError "Creating Dictionary Object"
         Exit Function
      End If
		Set wmi = GetObject("winmgmts:\\" & hostName & "\root\cimv2")
        'Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate,(Security)}!\\" & hostName & "\root\cimv2")
      If Err.Number <> 0 Then
         LogError "Creating WMI Object to connect to " & DQ(hostName)
         Exit Function
      End If
      '----------------------------------------------------------------------------------------------------------------------
      'Create the "SWbemDateTime" Object for converting WMI Date formats. Supported in Windows Server 2003 & Windows XP.
      '----------------------------------------------------------------------------------------------------------------------
      Set wmiDateTime = CreateObject("WbemScripting.SWbemDateTime")
      If Err.Number <> 0 Then
         LogError "Creating " & DQ("WbemScripting.SWbemDateTime") & " object"
         Exit Function
      End If
      '----------------------------------------------------------------------------------------------------------------------
      'Build the WQL query and execute it.
      '----------------------------------------------------------------------------------------------------------------------
      wmiDateTime.SetVarDate untilDateTime, True 'convert as local time
	  strUntil = wmiDateTime.Value
      wmiDateTime.SetVarDate fromDateTime, True 'convert as local time
	  strFrom = wmiDateTime.Value
      query = Left(query, InStrRev(query, "'")) & " And TimeWritten >= " & SQ(strFrom) & " And TimeWritten < " & SQ(strUntil) & ""
      'WScript.Echo "About to launch this query: " & query 
	  'query = "Select * from Win32_NTLogEvent Where Logfile = 'System' And TimeWritten >= '20140220095500.000000+000' And TimeWritten < '20140320095600.000000+000' "
      Set results = wmi.ExecQuery(query)
	  
      If Err.Number <> 0 Then
		 LogError  Err.Description
		 LogError  Err.Source
		 LogError  Err.Number
         LogError "Error Executing: " & DQ(query)
         Exit Function
      End If
	  'Wscript.Quit
	  
      For Each result In results
         Do
			With result
				'wmiDateTime.Value = .TimeWritten
				DateTime24 = ConvertWMIDateTime(.TimeWritten)
				cFullName = .ComputerName
				If Instr(cFullName,".")>0 Then 'trim off the domain name, if present, for the short name
					cShortName = Left(cFullName,Instr(cFullName,".")-1)
				Else
					cShortName = cFullName
				End If
					
					If filterExcludeEvents(excludeCodes,.EventCode) > 0 Then
						eventInfo=""
					Else
						eventInfo = DateTime24 _
						& "," & cShortName _
						& "," & cFullName _
						& "," & .EventCode _
						& "," & .LogFile _
						& "," & .Type _
						& "," & .SourceName _
						& "," & .CategoryString _
						& "," & .User _
						& "," & TrimEmpty(.Message)
	'							& "," & wmiDateTime.GetVarDate _
	'							& "," & .Category _
	'							& "," & .RecordNumber _
	
						Wscript.Echo eventInfo
					End If 
					
		
				
			End With
            If Err.Number <> 0 Then
               LogError "Enumerating Event Properties from the " & DQ(logName) & " event log on " & DQ(hostName)
				LogEror Err.Description
				LogError Err.Source
				LogError Err.Line
               errorCount = errorCount + 1
               Err.Clear
               Exit Do
            End If
			'WScript.Echo "About to add record " & CStr(eventsDict.Count+1) & " :: " &  eventinfo
			
			'If eventInfo <> "" Then 
			'	eventsDict.Item(CStr(eventsDict.Count+1)) = eventInfo
			'End If 
         Loop Until True
      Next
   'WScript.Echo "About to return a dictionary of " & CStr(eventsDict.Count) & " records. Error count is " & errorCount
   On Error Goto 0
   If errorCount <> 0 Then
      Exit Function
   End If
   QueryEventLog = True
End Function
'----------------------------------------------------------------------------------------------------------------------------
'Name       : NewDictionary -> Creates a new dictionary object.
'Parameters : None          ->
'Return     : NewDictionary -> Returns a dictionary object.
'----------------------------------------------------------------------------------------------------------------------------
Function NewDictionary
   Dim dict
   Set dict          = CreateObject("scripting.Dictionary")
   dict.CompareMode  = vbTextCompare
   Set NewDictionary = dict
End Function
'----------------------------------------------------------------------------------------------------------------------------
'Name       : SQ          -> Places single quotes around a string
'Parameters : stringValue -> String containing the value to place single quotes around
'Return     : SQ          -> Returns a single quoted string
'----------------------------------------------------------------------------------------------------------------------------
Function SQ(ByVal stringValue)
   If VarType(stringValue) = vbString Then
      SQ = "'" & stringValue & "'"
   End If
End Function
'----------------------------------------------------------------------------------------------------------------------------
'Name       : DQ          -> Place double quotes around a string and replace double quotes
'           :             -> within the string with pairs of double quotes.
'Parameters : stringValue -> String value to be double quoted
'Return     : DQ          -> Double quoted string.
'----------------------------------------------------------------------------------------------------------------------------
Function DQ (ByVal stringValue)
   If stringValue <> "" Then
	 
      DQ = """" & Replace (stringValue, """", """""") & """"
   Else
      DQ = """"""
   End If
End Function
'----------------------------------------------------------------------------------------------------------------------------
'Name      : LogMessage -> Parses a message to the log file.   
'Parameters: message    -> String containnig the message to include in the log file.
'Return    : None       ->    
'----------------------------------------------------------------------------------------------------------------------------
Function LogMessage(message)   
   If Not LogToFile(message) Then  
      Exit Function  
   End If  
End Function  
'----------------------------------------------------------------------------------------------------------------------------
'Name      : LogError -> Logs the current information about the error object.
'Parameters: message  -> String containnig the message that relates to the process that caused the error.
'Return    : None     ->    
'----------------------------------------------------------------------------------------------------------------------------
Function LogError(message)
   Dim errorMessage
   errorMessage = "Error " & Err.Number & " (Hex " & Hex(Err.Number) & ") " & message & ". " & Err.Description
   If Not LogToFile(errorMessage) Then
      Exit Function
   End If
End Function  
'----------------------------------------------------------------------------------------------------------------------------
'Name       : LogToFile -> Write a message into the user's network log file.   
'Parameters : LogSpec   -> String containing the Folder path, file name and extension of the log file to write to.   
'           : message   -> String containing the Message to be logged.   
'Return     : LogToFile -> Returns True if successful otherwise returns false.   
'----------------------------------------------------------------------------------------------------------------------------
Function LogToFile(message)
	WScript.echo message
   LogToFile = True  
End Function
'----------------------------------------------------------------------------------------------------------------------------
'Name      : BuildError -> Builds a string of information relating to the error object.
'Parameters: message    -> String containnig the message that relates to the process that caused the error.
'Return    : BuildError -> Returns a string relating to error object.   
'----------------------------------------------------------------------------------------------------------------------------
Function BuildError(message)
   BuildError = "Error " & Err.Number & " (Hex " & Hex(Err.Number) & ") " & message & ". " & Err.Description
End Function
'----------------------------------------------------------------------------------------------------------------------------
'Name       : ConvertWMIDateTime -> Converts a WMI Date Time String into a String that can be formatted as a valid Date Time.
'Parameters : wmiDateTimeString  -> String containing a WMI Date Time String.
'Return     : ConvertWMIDateTime -> Returns a valid Date Time String otherwise returns a Blank String.
'----------------------------------------------------------------------------------------------------------------------------
Function ConvertWMIDateTime(wmiDateTimeString)
	Dim wmiDateTime, integerValues, i
	Set wmiDateTime = CreateObject("WbemScripting.SWbemDateTime")
	If Err.Number <> 0 Then
		LogError "Creating " & DQ("WbemScripting.SWbemDateTime") & " object"
		Exit Function
	End If
	wmiDateTime.Value = wmiDateTimeString
	SetLocale(localeGB)
	ConvertWMIDateTime = FormatDateTime(wmiDateTime.GetVarDate(True),vbGeneralDate)
	SetLocale(localeHere)
End Function
'----------------------------------------------------------------------------------------------------------------------------
'Name       : TrimEmptyLines -> Find and remove empty lines in a string.
'Parameters : sString        -> aString passed by value
'Return     : TrimEmptyLines -> Returns a trimmed string
'----------------------------------------------------------------------------------------------------------------------------
Function TrimEmpty(sString)
 'On Error Resume Next
 'WScript.Echo "About to trim this string <" & aString & ">"
 If VarType(sString) <> vbString Then Exit Function
 Dim objRegExp
 Set objRegExp = New RegExp
' objRegExp.Pattern = "^\s+|\s+$|\r\n"
 objRegExp.Pattern = "\r\n\r\n|\r\n\s+|\s+\r\n"
 objRegExp.Ignorecase = True
 objRegExp.Global = True
 TrimEmpty = objRegExp.Replace(sString, vbCrLf)
 'TrimEmptyLines = "nothing "
End Function

'----------------------------------------------------------------------------------------------------------------------------
'Name       : QueryOSdetails -> Get the OS version, etc. Is needed to adjust the data processing
'Parameters : OS details     -> OS - OS version major, TZ - current TimeZone
'Return     : None           -> Results are returned via parameters
'----------------------------------------------------------------------------------------------------------------------------
Function QueryOSdetails(OS, TZ)
	Dim objWMIService, objItem, colItems
	Dim strComputer, strVersion

	On Error Resume Next
	strComputer = "."
	' WMI Connection to the object in the CIM namespace
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

	' WMI Query to the Win32_OperatingSystem
	Set colItems = objWMIService.ExecQuery("Select Version, CurrentTimeZone from Win32_OperatingSystem")

	' For Each... In Loop (Next at the very end)
	For Each objItem in colItems
		strVersion = Split(objItem.Version,".")
		OS = strVersion(0)
		TZ = objItem.CurrentTimeZone
		'WScript.Echo "Machine Name:     " & objItem.CSName 
		'WScript.Echo "Processor:        " & objItem.Description 
		'WScript.Echo "Manufacturer:     " & objItem.Manufacturer  
		'WScript.Echo "Operating System: " & objItem.Caption
		'WScript.Echo "Version:          " & objItem.Version
		'WScript.Echo "Service Pack:     " & objItem.CSDVersion 
		'WScript.Echo "CodeSet:          " & objItem.CodeSet  
		'WScript.Echo "CountryCode:      " & objItem.CountryCode 
		'WScript.Echo "OSLanguage:       " & objItem.OSLanguage  
		'WScript.Echo "CurrentTimeZone:  " & objItem.CurrentTimeZone 
		'WScript.Echo "Locale:           " & objItem.Locale
		'WScript.Echo "SerialNumber:     " & objItem.SerialNumber
		'WScript.Echo "SystemDrive:      " & objItem.SystemDrive
		'WScript.Echo "WindowsDirectory: " & objItem.WindowsDirectory
	Next
	'WScript.Echo "OS Version Major = " & OS & " and TZ = " & TZ
End Function

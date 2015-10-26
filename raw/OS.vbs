' OS.vbs

' -------------------------------------------------------' 
Option Explicit
Dim objWMIService, objItem, colItems
Dim strComputer, strList

On Error Resume Next
strComputer = "."

' WMI Connection to the object in the CIM namespace
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

' WMI Query to the Win32_OperatingSystem
Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")

' For Each... In Loop (Next at the very end)
For Each objItem in colItems
WScript.Echo "Machine Name:     " & objItem.CSName 
WScript.Echo "Processor:        " & objItem.Description 
WScript.Echo "Manufacturer:     " & objItem.Manufacturer  
WScript.Echo "Operating System: " & objItem.Caption
WScript.Echo "Version:          " & objItem.Version
WScript.Echo "Service Pack:     " & objItem.CSDVersion 
WScript.Echo "CodeSet:          " & objItem.CodeSet  
WScript.Echo "CountryCode:      " & objItem.CountryCode 
WScript.Echo "OSLanguage:       " & objItem.OSLanguage  
WScript.Echo "CurrentTimeZone:  " & objItem.CurrentTimeZone 
WScript.Echo "Locale:           " & objItem.Locale
WScript.Echo "SerialNumber:     " & objItem.SerialNumber
WScript.Echo "SystemDrive:      " & objItem.SystemDrive
WScript.Echo "WindowsDirectory: " & objItem.WindowsDirectory
Next

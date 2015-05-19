'###################################################################################################################################
'## This script was developed by Guberni and is part of Tellki's Monitoring Solution								  		      ##
'##																													  		      ##
'## December, 2014																									  		      ##
'##																													  		      ##
'## Version 1.0																										  		      ##
'##																													  		      ##
'## DESCRIPTION: Monitor disk space utilization																		  		      ##
'##																													  		      ##
'## SYNTAX: cscript "//Nologo" "//E:vbscript" "//T:90" "LogicalDisk.vbs" <HOST> <METRIC_STATE> <USERNAME> <PASSWORD> <DOMAIN>     ##
'##																													  		      ##
'## EXAMPLE: cscript "//Nologo" "//E:vbscript" "//T:90" "LogicalDisk.vbs" "10.10.10.1" "1,1,0" "user" "pwd" "domain"	  	      ##
'##																													              ##
'## README:	<METRIC_STATE> is generated internally by Tellki and its only used by Tellki default monitors. 						  ##
'##         1 - metric is on ; 0 - metric is off					              												  ##
'## 																												              ##
'## 	    <USERNAME>, <PASSWORD> and <DOMAIN> are only required if you want to monitor a remote server. If you want to use this ##
'##			script to monitor the local server where agent is installed, leave this parameters empty ("") but you still need to   ##
'##			pass them to the script.																						      ##
'## 																												              ##
'###################################################################################################################################

'Start Execution
Option Explicit
'Enable error handling
On Error Resume Next
If WScript.Arguments.Count <> 5 Then 
	CALL ShowError(3, 0)
End If
'Set Culture - en-us
SetLocale(1033)

'METRIC_ID
Const FreeSpace = "24:Free Space:4"
Const UsedSpace = "40:Used Space:4"
Const UsedSpacePerc = "11:% Used Space:6"


' INPUTS
Dim Host, MetricState, Username, Password, Domain
Host = WScript.Arguments(0)
MetricState = WScript.Arguments(1)
Username = WScript.Arguments(2)
Password = WScript.Arguments(3)
Domain = WScript.Arguments(4)


Dim arrMetrics
arrMetrics = Split(MetricState,",")
Dim objSWbemLocator, objSWbemServices, colItems
Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")

Dim Counter, Disks, objItem, FullUserName, Aux
Counter = 0
Disks = 0


	If Domain <> "" Then
		FullUserName = Domain & "\" & Username
	Else
		FullUserName = Username
	End If
	
	Set objSWbemServices = objSWbemLocator.ConnectServer(Host, "root\cimv2", FullUserName, Password)
	If Err.Number = -2147217308 Then
		Set objSWbemServices = objSWbemLocator.ConnectServer(Host, "root\cimv2", "", "")
		Err.Clear
	End If
	if Err.Number = -2147023174 Then
		CALL ShowError(4, Host)
		WScript.Quit (222)
	End If
	if Err.Number = -2147024891 Then
		CALL ShowError(2, Host)
	End If
	If Err Then CALL ShowError(1, Host)		
	
	if Err.Number = 0 Then
		if IsObject(objSWbemServices) = True Then
			objSWbemServices.Security_.ImpersonationLevel = 3
				Set colItems = objSWbemServices.ExecQuery("Select Name,Size,FreeSpace from Win32_LogicalDisk where Size is not null and DriveType=3",,16)
				If colItems.Count <> 0 Then
					For Each objItem in colItems
						'Free Space
						Disks = Disks + 1
						If arrMetrics(0)=1 Then _
						CALL Output(FreeSpace,FormatNumber(objItem.FreeSpace/1048576),objItem.Name)
						'Used Space
						Aux = (objItem.Size-objItem.FreeSpace)/1048576
						If arrMetrics(1)=1 Then _
						CALL Output(UsedSpace,FormatNumber(Aux),objItem.Name)
						'%Used Space
						If arrMetrics(2)=1 Then _
						CALL Output(UsedSpacePerc,FormatNumber((Aux/(objItem.Size/1048576))*100),objItem.Name)
					Next
				Else
					'If there is no response in WMI query
					CALL ShowError(5, Host)
				End If
			Rem End If
		End If
		If Err.number <> 0 Then
			CALL ShowError(5, Host)
			Err.Clear
		End If
	End If


If Err Then 
	CALL ShowError(1,0)
Else
	If Disks = 0 Then
		WScript.Quit(101)
	End If
	WScript.Quit(0)
End If

Sub ShowError(ErrorCode, Param)
	Dim Msg
	Msg = "(" & Err.Number & ") " & Err.Description
	If ErrorCode=2 Then Msg = "Access is denied"
	If ErrorCode=3 Then Msg = "Wrong number of parameters on execution"
	If ErrorCode=4 Then Msg = "The specified target cannot be accessed"
	If ErrorCode=5 Then Msg = "There is no response in WMI or returned query is empty"
	WScript.Echo Msg
	WScript.Quit(ErrorCode)
End Sub

Function GetOSVersion(SWbem)
	Dim colItems, objItem
	Set colItems = SWbem.ExecQuery("select BuildVersion from Win32_WMISetting",,16)
	For Each objItem in colItems
		GetOSVersion = CInt(objItem.BuildVersion)
	Next
End Function

Sub Output(MetricID, MetricValue, MetricObject)
	If MetricObject <> "" Then
		If MetricValue <> "" Then
			WScript.Echo MetricID & "|" & MetricValue & "|" & MetricObject & "|" 
		Else
			CALL ShowError(5, Host) 
		End If
	Else
		If MetricValue <> "" Then
			WScript.Echo MetricID & "|" & MetricValue & "|" 
		Else
			CALL ShowError(5, Host)
		End If
	End If
End Sub



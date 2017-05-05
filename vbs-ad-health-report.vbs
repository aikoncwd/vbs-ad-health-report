'###################################################
' 	TO-DO LIST
' 	[ ] - Telegram reports
'
'###################################################
'	IMPORTANT VARIABLES TO EDIT
'###################################################
	
	'This script will find automatically all your DC servers stored in the default OU 'Domain Controllers'
	'If you have a custom OU name for your DC server, set it on organizationUnitDC variable.
	'
	'If you have child domains or some complex Domain structure, just set usingOU = False and write
	'manually DNS names of every DC server you want to monitor/report (oDC Array variable)
	usingOU = True
	oDC = Array("SRV-DC1","SRV-DC2","SRV-AUX") 'Ignored if usingOU is True
	
	serviceState = "Running" 'Translate this into your System Language ONLY if the script is not working
	organizationUnitDC = "Domain Controllers" 'Default OU where DC's are stored, edit if you use a custom OU
	
	hardwareReport = True 'Additional table with hardware related information for each server
	minHDDfree = 30 'Minimun % HDD FreeSpace to mark as red/error
	minRAMfree = 20 'Minimun % RAM FreeSpace to mark as red/error
	
	includeRepadmin = True 'Include Replication Summary Report below the table. Nice to have, default: True
	
	emailReport = True 'Send the report using e-mail. Nice to have, default: True
	errorOnlyReport = False 'Send e-mail ONLY if an error occours
	attachCSV = True 'Attach a CSV version of the replication report/summary
	
	saveReport = True 'Save a file with the report summary
	pathReportOutput = "AD-health-summary.html" 'Path and name for the output file
	
	'E-mail settings, edit by your own
	emailSubject = "Active Directory Health Summary"
	emailFrom = "monitoring@mydomain.com"
	emailTo = "sysadmin@mydomain.com"
	emailCc = "" 
	emailPort = 25
	emailAuth = 1	
	emailSenderUser = "monitoring@mydomain.com"
	emailSenderPassword = "correcthorsebatterystaple"
	emailServer = "mail.mydomain.com"
	emailSSL = False
	
'###################################################
'	END OF IMPOTANT VARIABLES TO EDIT
'###################################################
'	BEGIN OF CODE - DON'T EDIT
'###################################################

GB = 1024 * 1024 * 1024
TimeStart = Timer()
isError = False
If usingOU Then
	Set oRD = GetObject("LDAP://RootDSE")
	Set oDC = GetObject("LDAP://ou=" & organizationUnitDC & ", " & oRD.Get("defaultNamingContext"))
End If
sHTML = ""
sHTML = sHTML & "<html>" & vbCrLf
sHTML = sHTML & "<head>" & vbCrLf
sHTML = sHTML & "<title>" & emailSubject & "</title>" & vbCrLf
sHTML = sHTML & "<style type='text/css'>" & vbCrLf
sHTML = sHTML & "<!--" & vbCrLf
sHTML = sHTML & "td {" & vbCrLf
sHTML = sHTML & "font-family: Tahoma;" & vbCrLf
sHTML = sHTML & "font-size: 11px;" & vbCrLf
sHTML = sHTML & "border-top: 1px solid #999999;" & vbCrLf
sHTML = sHTML & "border-right: 1px solid #999999;" & vbCrLf
sHTML = sHTML & "border-bottom: 1px solid #999999;" & vbCrLf
sHTML = sHTML & "border-left: 1px solid #999999;" & vbCrLf
sHTML = sHTML & "padding-top: 0px;" & vbCrLf
sHTML = sHTML & "padding-right: 0px;" & vbCrLf
sHTML = sHTML & "padding-bottom: 0px;" & vbCrLf
sHTML = sHTML & "padding-left: 0px;" & vbCrLf
sHTML = sHTML & "}" & vbCrLf
sHTML = sHTML & "body {" & vbCrLf
sHTML = sHTML & "margin-left: 3px;" & vbCrLf
sHTML = sHTML & "margin-top: 3px;" & vbCrLf
sHTML = sHTML & "margin-right: 3px;" & vbCrLf
sHTML = sHTML & "margin-bottom: 3px;" & vbCrLf
sHTML = sHTML & "table {" & vbCrLf
sHTML = sHTML & "border: thin solid #000000;" & vbCrLf
sHTML = sHTML & "}" & vbCrLf
sHTML = sHTML & "-->" & vbCrLf
sHTML = sHTML & "</style>" & vbCrLf
sHTML = sHTML & "</head>" & vbCrLf
sHTML = sHTML & "<body>" & vbCrLf

If hardwareReport Then
	sHTML = sHTML & "<table width='100%'>" & vbCrLf
	sHTML = sHTML & "<tr bgcolor='Lavender'>" & vbCrLf
	sHTML = sHTML & "<td colspan='7' height='25' align='center'>" & vbCrLf
	sHTML = sHTML & "<font face='tahoma' color='DarkBlue' size='4'><strong>Hardware Status Report</strong></font>" & vbCrLf
	sHTML = sHTML & "</td></tr></table>" & vbCrLf

	sHTML = sHTML & "<table width='100%'>" & vbCrLf
	sHTML = sHTML & "<tr bgcolor='FireBrick'>" & vbCrLf
	sHTML = sHTML & "<td align='center' rowspan=2><b><font color='white'>Server</font></b></td>" & vbCrLf
	sHTML = sHTML & "<td align='center' colspan=3><b><font color='white'>Disk Status</font></b></td>" & vbCrLf
	sHTML = sHTML & "<td align='center' colspan=3><b><font color='white'>RAM Status</font></b></td>" & vbCrLf
	sHTML = sHTML & "</tr>" & vbCrLf
	sHTML = sHTML & "<tr bgcolor='Indigo'>" & vbCrLf
	sHTML = sHTML & "<td align='center'><b><font color='white'>Total Size</font></b></td>" & vbCrLf
	sHTML = sHTML & "<td align='center'><b><font color='white'>Free Space</font></b></td>" & vbCrLf
	sHTML = sHTML & "<td align='center'><b><font color='white'>% Free</font></b></td>" & vbCrLf
	sHTML = sHTML & "<td align='center'><b><font color='white'>Total RAM</font></b></td>" & vbCrLf
	sHTML = sHTML & "<td align='center'><b><font color='white'>Free RAM</font></b></td>" & vbCrLf
	sHTML = sHTML & "<td align='center'><b><font color='white'>% Free</font></b></td>" & vbCrLf
	sHTML = sHTML & "</tr>" & vbCrLf

	If usingOU Then oDC.Filter = Array("Computer")

	For Each oComputer In oDC
		If usingOU Then
			remoteComputer = oComputer.CN
		Else
			remoteComputer = oComputer
		End If
		sHTML = sHTML & "<tr>" & vbCrLf
		sHTML = sHTML & "<td bgcolor='DarkGray' align=center><b>" & remoteComputer & "</b></td>" & vbCrLf
		Call checkDiskUsage(remoteComputer)
		Call checkRAMUsage(remoteComputer)
		sHTML = sHTML & "</tr>" & vbCrLf
	Next
	sHTML = sHTML & "</table><br>"
End If

sHTML = sHTML & "<table width='100%'>" & vbCrLf
sHTML = sHTML & "<tr bgcolor='Lavender'>" & vbCrLf
sHTML = sHTML & "<td colspan='7' height='25' align='center'>" & vbCrLf
sHTML = sHTML & "<font face='tahoma' color='DarkBlue' size='4'><strong>" & emailSubject & "</strong></font>" & vbCrLf
sHTML = sHTML & "</td></tr></table>" & vbCrLf
sHTML = sHTML & "<table width='100%'>" & vbCrLf
sHTML = sHTML & "<tr bgcolor='FireBrick'>" & vbCrLf
sHTML = sHTML & "<td align='center'><b><font color='white'>Server</font></b></td>" & vbCrLf
sHTML = sHTML & "<td align='center'><b><font color='white'>Ping Status</font></b></td>" & vbCrLf
sHTML = sHTML & "<td align='center'><b><font color='white'>Service DNS</font></b></td>" & vbCrLf
sHTML = sHTML & "<td align='center'><b><font color='white'>Service NTDS</font></b></td>" & vbCrLf
sHTML = sHTML & "<td align='center'><b><font color='white'>Service Netlogon</font></b></td>" & vbCrLf
sHTML = sHTML & "<td align='center'><b><font color='white'>Connectivity</font></b></td>" & vbCrLf
sHTML = sHTML & "<td align='center'><b><font color='white'>Advertising</font></b></td>" & vbCrLf
sHTML = sHTML & "<td align='center'><b><font color='white'>NetLogon</font></b></td>" & vbCrLf
sHTML = sHTML & "<td align='center'><b><font color='white'>Services</font></b></td>" & vbCrLf
sHTML = sHTML & "<td align='center'><b><font color='white'>Replication</font></b></td>" & vbCrLf
sHTML = sHTML & "<td align='center'><b><font color='white'>FSMO</font></b></td>" & vbCrLf
sHTML = sHTML & "<td align='center'><b><font color='white'>SysVol</font></b></td>" & vbCrLf
sHTML = sHTML & "<td align='center'><b><font color='white'>Topology</font></b></td>" & vbCrLf
sHTML = sHTML & "</tr>" & vbCrLf

If usingOU Then oDC.Filter = Array("Computer")

For Each oComputer In oDC
	If usingOU Then
		remoteComputer = oComputer.CN
	Else
		remoteComputer = oComputer
	End If
	sHTML = sHTML & "<tr>" & vbCrLf
	sHTML = sHTML & "<td bgcolor='DarkGray' align=center><b>" & remoteComputer & "</b></td>" & vbCrLf
    If checkPing(remoteComputer) Then
		sHTML = sHTML & "<td bgcolor='LightGreen' align=center><b>Success</b></td>" & vbCrLf
		Call checkService(remoteComputer, "DNS")
		Call checkService(remoteComputer, "NTDS")
		Call checkService(remoteComputer, "Netlogon")
		Call checkDcDiag(remoteComputer, "Connectivity")
		Call checkDcDiag(remoteComputer, "Advertising")
		Call checkDcDiag(remoteComputer, "NetLogons")
		Call checkDcDiag(remoteComputer, "Services")
		Call checkDcDiag(remoteComputer, "Replications")
		Call checkDcDiag(remoteComputer, "FsmoCheck")
		Call checkDcDiag(remoteComputer, "SysVolCheck")
		Call checkDcDiag(remoteComputer, "Topology")
	Else
		isError = True
		sHTML = sHTML & "<td bgcolor='Red' align=center><b>Failed</b></td>" & vbCrLf
		sHTML = sHTML & "<td bgcolor='Red' align=center><b>Failed</b></td>" & vbCrLf
		sHTML = sHTML & "<td bgcolor='Red' align=center><b>Failed</b></td>" & vbCrLf
		sHTML = sHTML & "<td bgcolor='Red' align=center><b>Failed</b></td>" & vbCrLf
		sHTML = sHTML & "<td bgcolor='Red' align=center><b>Failed</b></td>" & vbCrLf
		sHTML = sHTML & "<td bgcolor='Red' align=center><b>Failed</b></td>" & vbCrLf
		sHTML = sHTML & "<td bgcolor='Red' align=center><b>Failed</b></td>" & vbCrLf
		sHTML = sHTML & "<td bgcolor='Red' align=center><b>Failed</b></td>" & vbCrLf
		sHTML = sHTML & "<td bgcolor='Red' align=center><b>Failed</b></td>" & vbCrLf
		sHTML = sHTML & "<td bgcolor='Red' align=center><b>Failed</b></td>" & vbCrLf
		sHTML = sHTML & "<td bgcolor='Red' align=center><b>Failed</b></td>" & vbCrLf
		sHTML = sHTML & "<td bgcolor='Red' align=center><b>Failed</b></td>" & vbCrLf
	End If
	sHTML = sHTML & "</tr>" & vbCrLf
Next
sHTML = sHTML & "</table>"
If includeRepadmin Then sHTML = sHTML & checkRepadmin()
sHTML = sHTML & "<font face='tahoma' size='2'>Execution Time: <b>" & Round(Timer() - TimeStart, 2) & "s</b></font>"
sHTML = sHTML & "</body>"
sHTML = sHTML & "</html>"

If attachCSV Then
	Set oEXE = CreateObject("WScript.Shell").Exec("repadmin.exe /showrepl * /csv")
	Set F = CreateObject("Scripting.FileSystemObject").CreateTextFile("ad-health-summary.csv")
		F.Write oEXE.StdOut.ReadAll
	F.Close
End If

If emailReport Then Call sendMail(sHTML)

If saveReport Then
	Set F = CreateObject("Scripting.FileSystemObject").CreateTextFile(pathReportOutput)
		F.Write sHTML
	F.Close
End If

'###################################################
'	END OF CODE - DON'T EDIT
'###################################################
'	INTERNAL FUNCTIONS - DON'T EDIT
'###################################################

Function checkDiskUsage(RemoteComputer)
	Set oWMI = GetObject("winmgmts:\\" & RemoteComputer & "\root\CIMV2")
	Set colHDD = oWMI.ExecQuery("SELECT * FROM Win32_LogicalDisk WHERE Caption = 'C:'", "WQL", &H10 + &H20)
	For Each objHDD in colHDD
		HDDsize = Round(objHDD.Size / GB, 2)
		HDDfree = Round(objHDD.FreeSpace / GB, 2)
		HDDperc = Round((HDDfree * 100) / HDDsize, 2)
	Next
	If HDDperc < minHDDfree Then
		isError = True
		celColor = "Red"
	Else
		celColor = "LightGreen"
	End If
	sHTML = sHTML & "<td bgcolor='LightGreen' align=center><b>" & HDDsize & " Gb</b></td>" & vbCrLf
	sHTML = sHTML & "<td bgcolor='LightGreen' align=center><b>" & HDDfree & " Gb</b></td>" & vbCrLf
	sHTML = sHTML & "<td bgcolor='" & celColor & "' align=center><b>" & HDDperc & " %</b></td>" & vbCrLf
		
End Function

Function checkRAMUsage(RemoteComputer)
	Set oWMI = GetObject("winmgmts:\\" & RemoteComputer & "\root\CIMV2")
	Set colRAM = oWMI.ExecQuery("SELECT * FROM Win32_PerfFormattedData_PerfOS_Memory", "WQL", &H10 + &H20)
	For Each objRAM in colRAM
		RAMfree = Round(objRAM.AvailableBytes / GB, 2)
	Next
	Set colRAM = oWMI.ExecQuery("SELECT * FROM Win32_ComputerSystem", "WQL", &H10 + &H20)
	For Each objRAM in colRAM
		RAMsize = Round(objRAM.TotalPhysicalMemory / GB, 2)
	Next
	RAMperc = Round((RAMfree * 100) / RAMsize, 2)
	If RAMperc < minRAMfree Then
		isError = True
		celColor = "Red"
	Else
		celColor = "LightGreen"
	End If	
	sHTML = sHTML & "<td bgcolor='LightGreen' align=center><b>" & RAMsize & " Gb</b></td>" & vbCrLf
	sHTML = sHTML & "<td bgcolor='LightGreen' align=center><b>" & RAMfree & " Gb</b></td>" & vbCrLf
	sHTML = sHTML & "<td bgcolor='" & celColor & "' align=center><b>" & RAMperc & " %</b></td>" & vbCrLf
End Function

Function checkPing(RemoteComputer)
	Set oWMI = GetObject("winmgmts:\\.\root\CIMV2")
	If oWMI.Get("Win32_PingStatus.Address='" & RemoteComputer & "'").StatusCode = 0 Then
		checkPing = True
	Else
		checkPing = False
	End if
End Function

Function checkService(RemoteComputer, ServiceName)
	Set oWMI = GetObject("winmgmts:\\" & RemoteComputer & "\root\CIMV2")
	If oWMI.Get("Win32_Service.Name='" & serviceName & "'").State = serviceState Then
		sHTML = sHTML & "<td bgcolor='LightGreen' align=center><b>Success</b></td>" & vbCrLf
	Else
		isError = True
		sHTML = sHTML & "<td bgcolor='Red' align=center><b>Failed</b></td>" & vbCrLf
	End If
End Function

Function checkDcDiag(RemoteComputer, sTest)
	Set oEXE = CreateObject("WScript.Shell").Exec("dcdiag.exe /test:" & sTest & " /s:" & RemoteComputer)
	If InStr(oEXE.StdOut.ReadAll, "passed test " & sTest) > 0 Then
		sHTML = sHTML & "<td bgcolor='LightGreen' align=center><b>Success</b></td>" & vbCrLf
	Else
		isError = True
		sHTML = sHTML & "<td bgcolor='Red' align=center><b>Failed</b></td>" & vbCrLf
	End If
End Function

Function checkRepadmin()
	Set oEXE = CreateObject("WScript.Shell").Exec("repadmin.exe /replsummary * /bysrc /bydest")
	checkRepadmin = "<br><pre>" & oEXE.StdOut.ReadAll & "</pre><br>" & vbCrLf
End Function

Function sendMail(txtBody)
	If errorOnlyReport And isError Then Exit Function
	oCDO = "http://schemas.microsoft.com/cdo/configuration/"
	Set oMSG = CreateObject("CDO.Message")
		oMSG.Subject = emailSubject
		oMSG.From = emailFrom
		oMSG.To = emailTo
		If emailCc <> "" Then oMSG.Cc = emailCc
		oMSG.HTMLBody = txtBody
		If attachCSV Then oMSG.AddAttachment CreateObject("WScript.Shell").CurrentDirectory & "\ad-health-summary.csv"
		With oMSG.Configuration.Fields
			.Item(oCDO & "sendusing") = 2
			.Item(oCDO & "smtpserverport") = emailPort
			.Item(oCDO & "smtpauthenticate") = emailAuth
			.Item(oCDO & "smtpconnectiontimeout") = 60
			.Item(oCDO & "sendusername") = emailSenderUser
			.Item(oCDO & "sendpassword") = emailSenderPassword
			.Item(oCDO & "smtpserver") = emailServer
			.Item(oCDO & "smtpusessl") = emailSSL
			.Update
		End With
	oMSG.Send
End Function

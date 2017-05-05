'###################################################
' 	TO-DO LIST
' 	[ ] - Hardware report table
' 	[ ] - Telegram reports
'
'###################################################
'	IMPORTANT VARIABLES TO EDIT
'###################################################
	
	serviceState = "Running" 'Translate this into your default System Language
	organizationUnitDC = "Domain Controllers" 'Default OU where DC's are stored, edit if you a custom OU
	
	includeRepadmin = True 'Include the summary report below the table
	
	emailReport = True 'Send the report using e-mail
	attachCSV = True 'Attach a CSV version of the replication report/summary
	
	saveReport = True 'Save a file with the report summary
	pathReportOutput = "AD-health-summary.html" 'Path and name for the output file
	
	'E-mail settings, edit by your own
	emailSubject = "Active Directory Health Summary"
	emailFrom = "monitoring@mydomain.com"
	emailTo = "jgonzalez@cet10.com"
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

TimeStart = Timer()
Set oRD = GetObject("LDAP://RootDSE")
Set oDC = GetObject("LDAP://ou=" & organizationUnitDC & ", " & oRD.Get("defaultNamingContext"))

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

oDC.Filter = Array("Computer")
For Each oComputer In oDC
	sHTML = sHTML & "<tr>" & vbCrLf
	sHTML = sHTML & "<td bgcolor='DarkGray' align=center><b>" & oComputer.CN & "</b></td>" & vbCrLf
    If checkPing(oComputer.CN) Then
		sHTML = sHTML & "<td bgcolor='LightGreen' align=center><b>Success</b></td>" & vbCrLf
		Call checkService(oComputer.CN, "DNS")
		Call checkService(oComputer.CN, "NTDS")
		Call checkService(oComputer.CN, "Netlogon")
		Call checkDcDiag(oComputer.CN, "Connectivity")
		Call checkDcDiag(oComputer.CN, "Advertising")
		Call checkDcDiag(oComputer.CN, "NetLogons")
		Call checkDcDiag(oComputer.CN, "Services")
		Call checkDcDiag(oComputer.CN, "Replications")
		Call checkDcDiag(oComputer.CN, "FsmoCheck")
		Call checkDcDiag(oComputer.CN, "SysVolCheck")
		Call checkDcDiag(oComputer.CN, "Topology")
	Else
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
Next
sHTML = sHTML & "</tr>"
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
		sHTML = sHTML & "<td bgcolor='Red' align=center><b>Failed</b></td>" & vbCrLf
	End If
End Function

Function checkDcDiag(RemoteComputer, sTest)
	Set oEXE = CreateObject("WScript.Shell").Exec("dcdiag.exe /test:" & sTest & " /s:" & RemoteComputer)
	If InStr(oEXE.StdOut.ReadAll, "passed test " & sTest) > 0 Then
		sHTML = sHTML & "<td bgcolor='LightGreen' align=center><b>Success</b></td>" & vbCrLf
	Else
		sHTML = sHTML & "<td bgcolor='Red' align=center><b>Failed</b></td>" & vbCrLf
	End If
End Function

Function checkRepadmin()
	Set oEXE = CreateObject("WScript.Shell").Exec("repadmin.exe /replsummary * /bysrc /bydest")
	checkRepadmin = "<br><pre>" & oEXE.StdOut.ReadAll & "</pre><br>" & vbCrLf
End Function

Function sendMail(txtBody)
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

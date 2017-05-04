TimeStart = Timer()
Set oRD = GetObject("LDAP://RootDSE")
Set oDC = GetObject("LDAP://ou=Domain Controllers, " & oRD.Get("defaultNamingContext"))

sHTML = ""
sHTML = sHTML & "<html>" & vbCrLf
sHTML = sHTML & "<head>" & vbCrLf
sHTML = sHTML & "<title>Active Directory Health Summary</title>" & vbCrLf
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
sHTML = sHTML & "<font face='tahoma' color='DarkBlue' size='4'><strong>Active Directory Health Summary</strong></font>" & vbCrLf
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
    If CheckPing(oComputer.CN) Then
		sHTML = sHTML & "<td bgcolor='LightGreen' align=center><b>Success</b></td>" & vbCrLf
		Call CheckService(oComputer.CN, "DNS")
		Call CheckService(oComputer.CN, "NTDS")
		Call CheckService(oComputer.CN, "Netlogon")
		Call CheckDcDiag(oComputer.CN, "Connectivity")
		Call CheckDcDiag(oComputer.CN, "Advertising")
		Call CheckDcDiag(oComputer.CN, "NetLogons")
		Call CheckDcDiag(oComputer.CN, "Services")
		Call CheckDcDiag(oComputer.CN, "Replications")
		Call CheckDcDiag(oComputer.CN, "FsmoCheck")
		Call CheckDcDiag(oComputer.CN, "SysVolCheck")
		Call CheckDcDiag(oComputer.CN, "Topology")
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
sHTML = sHTML & "</table><br>"
sHTML = sHTML & CheckRepadmin()
sHTML = sHTML & "<br><font face='tahoma' size='2'>Execution Time: <b>" & Round(Timer() - TimeStart, 2) & "s</b></font>"
sHTML = sHTML & "</body>"
sHTML = sHTML & "</html>"

Call SendMail(sHTML)

Function CheckPing(RemoteComputer)
	Set oWMI = GetObject("winmgmts:\\.\root\CIMV2")
	If oWMI.Get("Win32_PingStatus.Address='" & RemoteComputer & "'").StatusCode = 0 Then
		CheckPing = True
	Else
		CheckPing = False
	End if
End Function

Function CheckService(RemoteComputer, ServiceName)
	Set oWMI = GetObject("winmgmts:\\" & RemoteComputer & "\root\CIMV2")
	If oWMI.Get("Win32_Service.Name='" & serviceName & "'").State = "Running" Then
		sHTML = sHTML & "<td bgcolor='LightGreen' align=center><b>Success</b></td>" & vbCrLf
	Else
		sHTML = sHTML & "<td bgcolor='Red' align=center><b>Failed</b></td>" & vbCrLf
	End If
End Function

Function CheckDcDiag(RemoteComputer, sTest)
	Set oEXE = CreateObject("WScript.Shell").Exec("dcdiag.exe /test:" & sTest & " /s:" & RemoteComputer)
	If InStr(oEXE.StdOut.ReadAll, "passed test " & sTest) > 0 Then
		sHTML = sHTML & "<td bgcolor='LightGreen' align=center><b>Success</b></td>" & vbCrLf
	Else
		sHTML = sHTML & "<td bgcolor='Red' align=center><b>Failed</b></td>" & vbCrLf
	End If
End Function

Function CheckRepadmin()
	Set oEXE = CreateObject("WScript.Shell").Exec("repadmin.exe /replsummary * /bysrc /bydest")
	CheckRepadmin = "<pre>" & oEXE.StdOut.ReadAll & "</pre>" & vbCrLf
End Function

Function SendMail(txtBody)
	oCDO = "http://schemas.microsoft.com/cdo/configuration/"
	Set oMSG = CreateObject("CDO.Message")
		oMSG.Subject = "Active Directory Health Summary"
		oMSG.From = "monitoring@mydomain.com"
		oMSG.To = "sysadmin@mydomain.com"
		'oMSG.Cc = "sysadmin2@mydomain.com" 
		oMSG.HTMLBody = txtBody
		With oMSG.Configuration.Fields
			.Item(oCDO & "sendusing") = 2
			.Item(oCDO & "smtpserverport") = 25
			.Item(oCDO & "smtpauthenticate") = 1
			.Item(oCDO & "smtpconnectiontimeout") = 60
			.Item(oCDO & "sendusername") = "monitoring@mydomain.com"
			.Item(oCDO & "sendpassword") = "p4ssw0rd"
			.Item(oCDO & "smtpserver") = "mail.monitoring.com"
			.Item(oCDO & "smtpusessl") = False
			.Update
		End With
	oMSG.Send
End Function
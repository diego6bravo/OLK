<!--#include file="chkLogin.asp" -->
<!--#include file="myHTMLEncode.asp"-->
<!--#include file="lang/autSubmit.asp" -->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252" >
<title>Untitled 1</title>
</head>
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<% 
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus


If Request("dbID") = "" Then 
	dbID = Session("ID") 
Else 
	dbID = CInt(Request("dbID"))
	UserName = Request("UserName")
End If

strAut = ""

arrAut = Split(Request("AutID"), ", ")

For i = 0 to UBound(arrAut)
	AutID = arrAut(i)
	
	If strAut <> "" Then strAut = strAut & "|"
	
	strAut = strAut & AutID & "%"
	
	If Request("chkAut" & AutID) = "Y" Then
		strAut = strAut & "{Y}"
	End If
	
	If Request("chkConf" & AutID) = "Y" Then
		strAut = strAut & "{C}"
	End If
	
	If Request("chkAutView" & AutID) = "Y" Then
		strAut = strAut & "{V}"
	End If
	
	If Request("Series" & AutID) <> "" and Request("Series" & AutID) <> "-1" Then
		strAut = strAut & "{S}" & Request("Series" & AutID)
	End If
	
	If Request("Series" & AutID & "2") <> "" and Request("Series" & AutID & "2") <> "-1" Then
		strAut = strAut & "{S2}" & Request("Series" & AutID & "2")
	End If

Next

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKSetUserAuth" & dbID
cmd.Parameters.Refresh
cmd("@Authorization") = strAut
cmd("@SlpCode") = Request("SlpCode")
If Request("Access") = "U" Then 
	cmd("@MaxDiscount") = Request("MaxDiscount") 
	cmd("@MaxDocDiscount") = Request("MaxDocDiscount") 
End If
cmd.execute()
%>
<body>
<script type="text/javascript">
<!--
alert('<%=getautSubmitLngStr("LtxtSaveDataConf")%>');
parent.enableCopy();
//-->
</script>
</body>

</html>

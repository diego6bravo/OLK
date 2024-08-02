<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="../chkLogin.asp" -->
<!--#include file="lang/sapUsers.asp" -->
<!--#include file="../myHTMLEncode.asp"-->
<html <% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>>

<head>
<%
set rx = Server.CreateObject("ADODB.recordset")
If Request("sapusers") <> "" Then
	arrVal = Split(Request("sapusers"), ", ")
	For i = 0 to UBound(arrVal)
		If i > 0 Then sqlV = sqlV & ", "
		sqlV = sqlV & "'" & Replace(arrVal(i), "'", "''") & "'"
	Next
End If
sql = "SELECT USER_CODE, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OUSR', 'U_NAME', T0.USERID, T0.U_NAME) U_NAME, "

If Request("sapusers") = "" Then
	sql = sql & " 'N' "
Else
	sql = sql & " Case When USER_CODE in (" & sqlV & ") Then 'Y' Else 'N' End "
End If

sql = sql & " Verfy FROM OUSR T0 WHERE Groups <> 99 order by U_Name"
rx.open sql, conn, 3, 1
%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />

<SCRIPT LANGUAGE="JavaScript">
function doAccept()
{
	var sapusers = document.form1.sapusers;
	selUCode = '';
	selUName = '';
	if (sapusers.length)
	{
		for (var i = 0;i<sapusers.length;i++)
		{
			if (sapusers[i].checked)
			{
				if (selUCode != '')
				{
					selUCode += ', ';
					selUName += ', ';
				}
				selUCode += sapusers[i].value.split('{|}')[1];
				selUName += sapusers[i].value.split('{|}')[0];
			}
		}
	}
	else
	{
		if (sapusers.checked)
		{
			selUCode = sapusers.value.split('{|}')[1];
			selUName = sapusers.value.split('{|}')[0];
		}
	}
	opener.sapTo(selUCode, selUName);
	window.close();
}

function check(field) {
	All = field.checked;
	sapusers = document.form1.sapusers;
	for (var i = 0;i<sapusers.length;i++)
	{
		sapusers[i].checked = All;
	}
}

function checkAll()
{
var All = true;
<% If rx.recordcount > 1 Then %>
var sapusers = document.form1.sapusers;
for (var i = 0;i<sapusers.length;i++)
{
	if (!sapusers[i].checked)
	{
		All = false;
		break;
	}
}
<% Else %>
if (!document.form1.sapusers.checked) { All = false; }
<% End If %>
if (document.form1.C1 != null) document.form1.C1.checked = All;
}
</script>
<link rel="stylesheet" type="text/css" href="../design/0/style/stylePopUp.css">
<title><%=getsapUsersLngStr("LttlSAPUsers")%></title>
</head>

<body topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" onload="checkAll();">
<form method="post" action="sapusers.asp" name="form1">
            <table border="0" cellpadding="0"  bordercolor="#111111" width="100%" id="table1">
              <tr class="GeneralTlt">
                <td width="50%" class="GeneralTlt"><%=getsapUsersLngStr("LttlSAPUsers")%>:</td>
              </tr>
              <tr>
                <td width="50%">
                <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="table2">
              <% do while not rx.eof %>
                  <tr class="GeneralTbl">
                    <td width="100%">
					<input type="checkbox" <% If rx("Verfy") = "Y" Then %>checked<% End If %> name="sapusers" onclick="checkAll()" value="<%=myHTMLEncode(RX("U_Name"))%>{|}<%=RX("User_Code")%>" id="fps<%=RX("User_Code")%>" <%=chkItem(RX("User_Code"))%> style="border-style:solid; border-width:0; background:background-image"><label for="fps<%=RX("User_Code")%>"><%=RX("U_Name")%></label>
					</td>
                  </tr>
		        <% rx.movenext
		        loop %>
		        <% If rx.recordcount > 1 Then %>
                  <tr class="GeneralTbl">
                    <td width="100%" height="21">
					<input type="checkbox" <% If UBound(Split(Request.QueryString("sapusers"),", ")) = rx.recordcount-1 then Response.Write "checked" %> name="C1" value="ON" onclick="check(this)" id="fp1" style="border-style:solid; border-width:0; background:background-image"><label for="fp1"><%=getsapUsersLngStr("DtxtAll")%></label></td>
                  </tr>
                <% End If %>
                  <tr class="GeneralTbl">
                    <td width="100%" height="21">
					<p align="center">
					<input type="button" value="<%=getsapUsersLngStr("DtxtAccept")%>" name="B1" onclick="javascript:doAccept();"></td>
                  </tr>
                </table>
                </td>
              </tr>
              </table>
			<input type="hidden" name="AddPath" value="../">
			<input type="hidden" name="pop" value="Y">
</body>

</html>
<% conn.close
set rx = nothing 
set rx = nothing
Public Function chkItem(varx)
	ArrVal = Split(Request.QueryString("sapusers"),", ")
	For i = 0 to UBound(ArrVal)
		if ArrVal(i) = varx then chkItem = "checked"
	next
End Function
%>
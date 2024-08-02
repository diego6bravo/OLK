<!--#include file="chkLogin.asp" -->
<!--#include file="myHTMLEncode.asp" -->
<!--#include file="lang/adminEditDec.asp" -->
<html <% If Session("rtl") <> "" Then %>dir="rtl"<% End If %>>

<head>
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

%>

<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=getadminEditDecLngStr("LtxtDesProp")%></title>
<script language="javascript" src="general.js"></script>
<link rel="stylesheet" href="style/style_pop.css">
</head>

<body topmargin="0" leftmargin="0" bgcolor="#F5FBFE" onbeforeunload="javascript:opener.clearWin();">
<% conn.execute("use OLKCommon")
If Request.Form.Count > 0 Then
sql = "declare @DisID int set @DisID = " & Request("ID") & " " & _
"declare @AlterName nvarchar(50) set @AlterName = N'" & saveHTMLDecode(Request("AlterName"), False) & "' " & _
"declare @Note nvarchar(256) set @Note = N'" & saveHTMLDecode(Request("Note"), False) & "' " & _
"if @Note = '' begin set @Note = null end " & _
"if not exists(select 'A' from OLKCustDes where DisID = @DisID) begin " & _
"	insert OLKCustDes (DisID, AlterName, Note) values(@DisID, @AlterName, @Note) " & _
"end else begin " & _
"	update OLKCustDes set AlterName = @AlterName, Note = @Note where DisID = @DisID " & _
"end"
conn.execute(sql) %>
<script language="javascript">
window.close();
opener.setDisName(<%=Request("ID")%>, '<%=Replace(Request("AlterName"), "'", "\'")%>');
</script>
<% Else   
set rs = Server.CreateObject("ADODB.RecordSet")

sql = "select * from OLKCustDes where DisID = " & Request("ID")
set rs = conn.execute(sql)

If Not rs.Eof Then
	AlterName = rs("AlterName")
	Note = rs("Note")
Else
	AlterName = Request("Name")
End If
%>
<script language="javascript">
function valFrm()
{
	if (document.frm.AlterName.value == '')
	{
		alert('<%=getadminEditDecLngStr("LtxtValAlterNam")%>');
		document.frm.AlterName.focus();
		return false;
	}
	return true;
}
</script>
<table border="0" width="100%" id="table1" cellpadding="0">
	<form name="frm" method="POST" action="adminEditDec.asp" onsubmit="return valFrm();" webbot-action="--WEBBOT-SELF--">
	<tr>
		<td class="popupTtl" colspan="2"><%=getadminEditDecLngStr("LtxtDesProp")%> - <%=Server.HTMLEncode(Request("Name"))%></td>
	</tr>
	<tr>
		<td class="popupOptDesc" style="width: 100px"><%=getadminEditDecLngStr("DtxtAlertName")%></td>
		<td class="popupOptValue">
		<input type="text" name="AlterName" size="20" class="input" style="width:100%" maxlength="100" value="<%=Server.HTMLEncode(AlterName)%>"></td>
	</tr>
	<tr>
		<td class="popupOptDesc" style="width: 100px" valign="top"><%=getadminEditDecLngStr("DtxtNote")%></td>
		<td class="popupOptValue">
		<textarea rows="4" name="Note" cols="20" class="input" style="width:100%; height: 56px; " maxlength="256"><% If Not IsNull(Note) Then %><%=Server.HTMLEncode(Note)%><% End If %></textarea></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE" colspan="2">
		<table border="0" width="100%" id="table2" cellpadding="0">
			<tr>
				<td style="width: 75px"><font color="#4783C5" face="Verdana" size="1">
		<input type="submit" value="<%=getadminEditDecLngStr("DtxtSave")%>" name="btnSave" class="OlkBtn"></font></td>
				<td>
								<hr color="#0D85C6" size="1"></td>
				<td style="width: 75px">
				<p align="right"><font color="#4783C5" face="Verdana" size="1">
				<input type="button" value="<%=getadminEditDecLngStr("DtxtCancel")%>" name="btnCancel" class="OlkBtn" onclick="window.close();"></font></td>
			</tr>
		</table>
		</td>
	</tr>
	<input type="hidden" name="ID" value="<%=Request("ID")%>">
	</form>
</table>
<% End If %>

</body>
<%
set rs = nothing
conn.close 
%>
</html>
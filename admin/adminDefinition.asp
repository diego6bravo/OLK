<html <% If Session("rtl") <> "" Then %>dir="rtl" <% End If %>>
<!--#include file="chkLogin.asp" -->
<!--#include file="myHTMLEncode.asp"-->
<!--#include file="lang/adminDefinition.asp" -->
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<% 
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus


If Session("ID") = "" Then Response.Redirect "default.asp"

conn.execute("use [" & Session("olkdb") & "]") 

set rs = Server.CreateObject("ADODB.RecordSet")

If Request("IsNew") <> "Y" Then
	sql = "select Definition " & _
			"from OLKQryDefinitions T0 " & _
			"where PageID = " & Request.Form("PageID") & " and FieldID = '" & Request.Form("FieldID") & "' and FieldKey = '" & Request.Form("FieldKey") & "'"
	set rs = conn.execute(sql)
	If Not rs.Eof Then txtDefinition = rs(0)
Else
	txtDefinition = Request("NewValue")
	If txtDefinition <> "" Then
		txtDefinition = Split(txtDefinition, "{S}")(2)
	End If
End If

%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-<%=getadminDefinitionLngStr("charset")%>" />
<title><%=getadminDefinitionLngStr("LttlDefinition")%></title>
<style type="text/css">
body{
scrollbar-base-color:#336699;
scrollbar-highlight-color:#E1F1FF;
}

</style>
<link rel="stylesheet" type="text/css" href="style/style_pop.css"/>
</head>

<body style="margin: 0px; " onbeforeunload="opener.clearWin();">

<form method="post" action="<% If Request("IsNew") <> "Y" Then %>adminSubmit.asp<% Else %>adminDefinitionReturn.asp<% End If %>">
	<table style="width: 100%" cellpadding="0">
		<tr class="popupTtl">
			<td colspan="2">&nbsp;<%=getadminDefinitionLngStr("LttlDefinition")%></td>
		</tr>
		<tr class="TblGreenNrm">
			<td colspan="2" align="center">
			<textarea name="txtDefinition" style="width: 98%; height: 194px;"><%=txtDefinition%></textarea>
			</td>
		</tr>
		<tr>
			<td colspan="2">
			<div align="center">
				<table border="0" cellpadding="0" id="table6" style="width: 100%">
					<tr>
						<td style="width: 75px">
						<p align="center">
						<input type="submit" value="<%=getadminDefinitionLngStr("DtxtAccept")%>" name="btnAccept" class="OlkBtn" /></p>
						</td>
						<td >
							<hr color="#0D85C6" size="1"/></td>
						<td style="width: 75px">
						<p align="center">
						<input type="button" value="<%=getadminDefinitionLngStr("DtxtCancel")%>" name="btnCancel" class="OlkBtn" onClick="javascript:window.close();" /></p></td>
					</tr>
				</table>
			</div>
			</td>
		</tr>
	</table>
	<input type="hidden" name="PageID" value='<%=Request("PageID")%>' />
	<input type="hidden" name="FieldID" value='<%=Request("FieldID")%>' />
	<input type="hidden" name="FieldKey" value='<%=Request("FieldKey")%>' />
	<input type="hidden" name="pop" value="Y" />
	<input type="hidden" name="submitCmd" value="adminDefinition" />
	<input type="hidden" name="new" value='<%=Request("new")%>' />
</form>

</body>

</html>

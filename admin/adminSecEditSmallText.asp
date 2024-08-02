<!--#include file="chkLogin.asp" -->
<!-- #INCLUDE file="FCKeditor/fckeditor.asp" -->
<!--#include file="myHTMLEncode.asp" -->
<!--#include file="lang/adminSecEditSmallText.asp" -->
<html <% If Session("rtl") <> "" Then %>dir="rtl"<% End If %>>
<link href="style/style_pop.css" rel="stylesheet" type="text/css">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=getadminSecEditSmallTextLngStr("LttlEditShortText")%></title>
<script language="javascript" src="general.js"></script>
</head>

<script language="javascript">
<% If Request.Form("btnAccept") <> "" Then %>
opener.setSmallText('<%=Replace(Replace(myHTMLEncode(Request("smallText")), VbNewLine, "\n"), "'", "\'")%>');
window.close();
<% End If %>
</script>
<body topmargin="0" leftmargin="0" style="background-color: #F7FBFF" onbeforeunload="javascript:opener.clearWin();">
<form method="POST" name="frmSmallText" action="adminSecEditSmallText.asp" webbot-action="--WEBBOT-SELF--">
	<%
	Dim oFCKeditor
	Set oFCKeditor = New FCKeditor
	oFCKeditor.BasePath = "FCKeditor/"
	oFCKeditor.Height = 270
	oFCKEditor.ToolbarSet = "Custom"
	oFCKEditor.Value = Request.Form("smallText")
	oFCKEditor.Config("AutoDetectLanguage") = False
	If Session("myLng") <> "pt" Then
		oFCKEditor.Config("DefaultLanguage") = Session("myLng")
	Else
		oFCKEditor.Config("DefaultLanguage") = "pt-br"
	End If
	oFCKeditor.Create "smallText"
	%>
<table border="0" cellpadding="0" cellspacing="1" width="100%" id="table3">
	<tr>
		<td width="75" id="tdSave">
<input type="submit" value="<%=getadminSecEditSmallTextLngStr("DtxtAccept")%>" name="btnAccept" class="OlkBtn"></td>
		<td>
<hr size="1">
		</td>
		<td width="75">
		<p align="right">
<input type="button" value="<%=getadminSecEditSmallTextLngStr("DtxtCancel")%>" name="B1" class="OlkBtn" onClick="javascript:if(confirm('<%=getadminSecEditSmallTextLngStr("LtxtValCloseWin")%>'))window.close();"></td>
	</tr>
</table>
</form>
</body>
</html>
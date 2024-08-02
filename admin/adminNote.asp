<!--#include file="chkLogin.asp" -->
<!--#include file="myHTMLEncode.asp"-->
<!--#include file="lang/adminNote.asp" -->
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

           
If Request("cmd") = "e" then 
      GetQuery rs, 6, "Y", Request("noteIndex")
      noteVar = RS("Note")
      noteName = Server.HTMLEncode(RS("NoteName"))
      set rs = Nothing
      conn.close
End If
%>
<html <% If Session("rtl") <> "" Then %>dir="rtl"<% End If %>>
<link href="style/style_pop.css" rel="stylesheet" type="text/css">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><% If Request("cmd") = "e" then %><%=getadminNoteLngStr("LttlEditNote")%><% elseif Request("cmd") = "a" then %><%=getadminNoteLngStr("LttlAddNote")%><% end if %></title>
<script language="javascript" src="general.js"></script>
<style type="text/css">
.style1 {
	font-size: xx-small;
	color: #31659C;
}
</style>
</head>

<body topmargin="0" leftmargin="0" onbeforeunload="opener.clearWin();">
<script language="javascript">
function valFrm()
{
	if (document.form1.NoteName.value == '')
	{
		alert('<%=getadminNoteLngStr("LtxtValFldNam")%>');
		return false;
	}
	else if (document.form1.NoteName.value.length > 50)
	{
		alert('<%=getadminNoteLngStr("LtxtValFldLength")%>');
		return false;
	}
	else if (document.form1.Note.value == '')
	{
		alert('<%=getadminNoteLngStr("LtxtValNote")%>');
		return false;
	}
	else if (document.form1.Note.value.length > 254)
	{
		alert('<%=getadminNoteLngStr("LtxtValFldMaxCharNote")%>');
		return false;
	}
	return true;
}
</script>
<form method="POST" action="adminSubmit.asp" name="form1" onsubmit="return valFrm();">
	<table border="0" cellpadding="0" width="100%" height="110">
		<tr>
			<td valign="top" bgcolor="#FFFFFF">
			<div align="left">
				<table border="0" cellpadding="0" id="table4" style="width: 100%">
					<tr>
						<td class="popupTtl"><% If Request("cmd") = "e" then %><%=getadminNoteLngStr("LttlEditNote")%><% elseif Request("cmd") = "a" then %><%=getadminNoteLngStr("LttlAddNote")%><% end if %></td>
					</tr>
					<tr>
						<td>
						<table border="0" cellpadding="0" width="100%" id="table5" height="63">
							<tr>
								<td  class="popupOptDesc" style="width: 100px">
								<%=getadminNoteLngStr("DtxtName")%></td>
								<td class="popupOptValue">
								<input type="text" name="NoteName" size="45" maxlength="50" class="input" style="width: 100%" value="<%=NoteName%>"></td>
							</tr>
							<tr>
								<td class="popupOptDesc" style="width: 100px" valign="top">
								<%=getadminNoteLngStr("DtxtNote")%></td>
								<td class="popupOptValue">
    							<textarea rows="5" name="Note" style="width: 100%; height: 57px;" class="input" onfocus="javascript:this.select()" maxlength="254"><% If Not IsNull(NoteVar) Then %><%=Server.HTMLEncode(NoteVar)%><% End If %></textarea></td>
							</tr>
						</table>
						</td>
					</tr>
					<tr>
						<td>
						<table border="0" cellpadding="0" width="100%" id="table6">
							<tr>
								<td style="width: 75px">
								<p align="center">
								<input type="submit" value="<%=getadminNoteLngStr("DtxtConfirm")%>" name="B2" class="OlkBtn"></td>
								<td>
								<hr color="#0D85C6" size="1"></td>
								<td style="width: 75px">
								<p align="center">
								<input type="button" value="<%=getadminNoteLngStr("DtxtCancel")%>" name="B1" class="OlkBtn" onClick="javascript:window.close();"></td>
							</tr>
						</table>
						</td>
					</tr>
				</table>
			</div>
			</td>
		</tr>
	</table>
	<input type="hidden" name="cmd" value="<%=Request("cmd")%>">
	<input type="hidden" name="noteIndex" value="<%=Request("NoteIndex")%>">
	<input type="hidden" name="submitCmd" value="adminnote">
	<input type="hidden" name="LineNum" value="<%=Request("LineNum")%>">
</form>

</body>

</html>
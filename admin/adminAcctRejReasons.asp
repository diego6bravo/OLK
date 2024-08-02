<!--#include file="chkLogin.asp" -->
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->

<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
           
If Request("cmd") = "e" then 
      sql = "select ReasonName, Reason from OLKAcctRejectNotes where ReasonIndex = " & Request("rIndex")
      set rs = conn.execute(sql) 
      Reason = RS("Reason")
      ReasonName = RS("ReasonName")
      set rs = Nothing
Else
	Reason = ""
	ResonName = ""
End If
conn.close
%>
<!--#include file="myHTMLEncode.asp"-->
<!--#include file="lang/adminAcctRejReasons.asp" -->
<html <% If Session("rtl") <> "" Then %>dir="rtl"<% End If %>>
<link rel="stylesheet" type="text/css" href="style/style_pop.css"/>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><% If Request("cmd") = "e" then %><%=getadminAcctRejReasonsLngStr("DtxtEdit")%><% elseif Request("cmd") = "a" then %><%=getadminAcctRejReasonsLngStr("DtxtAdd")%><% end if %>&nbsp;<%=getadminAcctRejReasonsLngStr("LtxtReason")%></title>
</head>
<script language="javascript" src="general.js"></script>
<script language="javascript">
<!--
function chkMax(e, f, m)
{
	if(f.value.length == m && (e.keyCode != 8 && e.keyCode != 9 && e.keyCode != 35 && e.keyCode != 36 && e.keyCode != 37 
	&& e.keyCode != 38 && e.keyCode != 39 && e.keyCode != 40 && e.keyCode != 46 && e.keyCode != 16))return false; else return true;
}

function valFrm()
{
	if (document.frmReason.ReasonName.value == '') {
		alert('<%=getadminAcctRejReasonsLngStr("LtxtValNamReason")%>');
		document.frmReason.ReasonName.focus();
		return false;
	}
	else if (document.frmReason.Reason.value == '') {
		alert('<%=getadminAcctRejReasonsLngStr("LtxtValReason")%>');
		document.frmReason.Reason.focus();
		return false;
	}
	return true;
}
//-->
</script>
<body topmargin="0" leftmargin="0" onbeforeunload="opener.clearWin();">
<form method="POST" action="adminSubmit.asp" name="frmReason" onsubmit="return valFrm();">
	<table border="0" cellpadding="0" width="100%" height="110" id="table1">
		<tr>
			<td valign="top">
			<div align="left">
				<table border="0" cellpadding="0" width="100%" id="table4">
					<tr>
						<td class="popupTtl"><% If Request("cmd") = "e" then %><%=getadminAcctRejReasonsLngStr("DtxtEdit")%><% elseif Request("cmd") = "a" then %><%=getadminAcctRejReasonsLngStr("DtxtAdd")%><% end if %>&nbsp;<%=getadminAcctRejReasonsLngStr("LtxtReason")%></td>
					</tr>
					<tr>
						<td>
						<table border="0" cellpadding="0" width="100%" id="table5">
							<tr>
								<td class="popupOptDesc" width="100"><%=getadminAcctRejReasonsLngStr("DtxtName")%></td>
								<td class="popupOptValue">
    							<input type="text" name="ReasonName" size="45" style="width: 259px;" class="input" value="<%=Server.HTMLEncode(ReasonName)%>" onkeydown="return chkMax(event, this, 50);"></td>
							</tr>
							<tr>
								<td class="popupOptDesc" colspan="2">
								<p align="center"><%=getadminAcctRejReasonsLngStr("LtxtReason")%></td>
							</tr>
							<tr>
								<td class="popupOptValue" colspan="2">
								<textarea name="Reason" class="input" onfocus="javascript:this.select()" style="width: 100%; height: 206px;"><%=Server.HTMLEncode(Reason)%></textarea></td>
							</tr>
						</table>
						</td>
					</tr>
					<tr>
						<td>
						<div align="center">
							<table border="0" cellpadding="0" id="table6" style="width: 100%">
								<tr>
									<td style="width: 75px">
									<p align="center">
									<input type="submit" value="<%=getadminAcctRejReasonsLngStr("DtxtConfirm")%>" name="B2" class="OlkBtn">
									</td>
									<td >
								<hr color="#0D85C6" size="1"></td>
									<td style="width: 75px">
									<p align="center">
									<input type="button" value="<%=getadminAcctRejReasonsLngStr("DtxtCancel")%>" name="B1" class="OlkBtn" onClick="javascript:window.close();"></td>
								</tr>
							</table>
						</div>
						</td>
					</tr>
				</table>
			</div>
			</td>
		</tr>
	</table>
	<input type="hidden" name="cmd" value="<%=Request("cmd")%>">
	<input type="hidden" name="submitCmd" value="adminAcctRejReasons">
	<input type="hidden" name="rIndex" value="<%=Request("rIndex")%>">
</form>

</body>

</html>
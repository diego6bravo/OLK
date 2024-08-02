<!--#include file="clientInc.asp"-->
<!--#include file="agentTop.asp"-->
<% If userType <> "V" Then Response.Redirect "default.asp" %>
<% addLngPathStr = "" %>
<!--#include file="lang/changePwd.asp" -->
<% 
Verfy = True
If Request.Form("btnSave") <> "" Then
	sql = "select case when (select Password from OLKAgentsAccess where SlpCode = " & Session("vendid") & ") = N'" & Request("txtpass1") & "' Then 'Y' Else 'N' End Verfy"
	set rs = conn.execute(sql)
	Verfy = rs("Verfy") = "Y"
	
	If Verfy Then
		sql = "update OLKAgentsAccess set Password = N'" & Request("txtpass2") & "' where SlpCode = " & Session("vendid")
		conn.execute(sql)
	End If
End If 

If Verfy and Request("btnSave") <> "" Then %>
<table border="0" width="100%" cellpadding="0" id="table1">
	<tr>
		<td class="TablasTituloSec" id="tdMyTtl">&nbsp;<%=getchangePwdLngStr("LtxtChangePwd")%></td>
	</tr>
	<tr>
		<td class="FirmTbl">
		<p align="center"><%=getchangePwdLngStr("LtxtOKUpd")%></td>
	</tr>
</table>
<% Else %>
<script language="javascript">
<!--
<% If Not Verfy Then %>
alert('<%=getchangePwdLngStr("LtxtErrCurPwd")%>');
<% End If %>
function verDatos(){
var strPass1 = window.document.frmData.txtpass1.value 
var strPass2 = window.document.frmData.txtpass2.value 
var strPass3 = window.document.frmData.txtpass3.value    

if (isEmpty(strPass1)){
	alert("<%=getchangePwdLngStr("LtxtValCurPwd")%>")
	window.document.frmData.txtpass1.focus()
	return false;  
	}
	
if (isEmpty(strPass2)){
	alert("<%=getchangePwdLngStr("LtxtValNewPwd")%>")
	window.document.frmData.txtpass2.focus()
	return false;  
	}
	
if (isEmpty(strPass3)){
	alert("<%=getchangePwdLngStr("LtxtValConfPwd")%>")
	window.document.frmData.txtpass3.focus()
	return false;  
	}
	
if (!verPwd(strPass2,strPass3)){
	return false;
}

return true;
}

function verPwd(pPwd1,pPwd2) {
if (pPwd1 != pPwd2){
	alert("<%=getchangePwdLngStr("LtxtValNoMatchNewPwd")%>")
	window.document.frmData.txtpass2.value = '';
	window.document.frmData.txtpass3.value = '';
	document.frmData.txtpass2.focus();
	return false
	}
return true
}


function isEmpty(inputVal) {
inputStr = inputVal.length
var contsps = 0    // contador de espacios en blanco
for (var i = 0; i < inputStr ; i++) {
	var oneChar = inputVal.charAt(i)
	if (oneChar == " ") {
	contsps = contsps + 1 
	}
}
if (contsps == inputStr) {
	return true
 }else {
        return false }
}


//-->
</script> 
<form method="POST" action="changePwd.asp" name="frmData" onsubmit="javascript:return verDatos();">
<input type="hidden" name="cmd" value="<%=Request("cmd")%>">
<table border="0" cellpadding="0" width="100%" id="table1">
	<tr>
		<td class="TablasTituloSec" id="tdMyTtl">&nbsp;<%=getchangePwdLngStr("LtxtChangePwd")%></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table2">
			<tr>
				<td width="137" class="DatosTltIn"><%=getchangePwdLngStr("LtxtCurPwd")%></td>
				<td class="TblGeneral">
				<input type="password" name="txtpass1" size="31" onkeydown="return chkMax(event, this, 20);" maxlength="20"></td>
			</tr>
			<tr>
				<td width="137" class="DatosTltIn"><%=getchangePwdLngStr("LtxtNewPwd")%></td>
				<td class="TblGeneral">
				<input type="password" name="txtpass2" size="31" onkeydown="return chkMax(event, this, 20);" maxlength="20"> 
				5 - 20 <%=getchangePwdLngStr("LtxtChars")%></td>
			</tr>
			<tr>
				<td width="137" class="DatosTltIn"><%=getchangePwdLngStr("LtxtConfPwd")%></td>
				<td class="TblGeneral">
				<input type="password" name="txtpass3" size="31" onkeydown="return chkMax(event, this, 20);" maxlength="20"></td>
			</tr>
			<tr>
				<td class="DatosTltIn" align="right" colspan="2">
				<p align="center">
				<input type="submit" class="BtnSave" value="<%=getchangePwdLngStr("DtxtSave")%>" name="btnSave"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
</table>
</form>
<% End If %>
<!--#include file="agentBottom.asp"-->
<%@ Language=VBScript %>
<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="../chkLogin.asp" -->
<!--#include file="../clearItem.asp"-->
<!--#include file="../myHTMLEncode.asp"-->
<!--#include file="lang/AddCartGetTaxCode.asp" -->
<html <% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=getAddCartGetTaxCodeLngStr("LtxtVatGrp")%></title>
<link rel="stylesheet" type="text/css" href="../design/0/style/stylePopUp.css">
</head>

<body topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0">
<script language="javascript" src="../general.js"></script>
<%
If Request.Form("btnConfirm") = "" Then
set rs = Server.CreateObject("ADODB.recordset")
If Request("expItem") <> "Y" Then
	sql = "select ItemCode, IsNull(ItemName, '') ItemName from oitm T0 where T0.ItemCode = N'" & saveHTMLDecode(Request("Item"), False) & "'"
Else
	sql = "SELECT ExpnsCode ItemCode, IsNull(ExpnsName, '') ItemName FROM OEXD T0 where ExpnsCode = " & Request("Item")
End If
set rs = conn.execute(sql)
%>
<script language="javascript">
function valFrm()
{
	if (document.frmTax.TaxCode.selectedIndex == 0)
	{
		alert("<%=getAddCartGetTaxCodeLngStr("LtxtValSlVat")%>");
		document.frmTax.TaxCode.focus();
		return false;
	}
	return true;
}
</script>
<form method="POST" action="AddCartGetTaxCode.asp" name="frmTax" onsubmit="return valFrm();">
<div align="left">
	<table border="0" cellpadding="0" width="300" id="table1">
		<tr class="GeneralTbl">
			<td><%=rs("ItemCode")%>&nbsp;</td>
			<td><%=rs("ItemName")%>&nbsp;</td>
		</tr>
		<tr class="GeneralTbl">
			<td colspan="2"><select size="1" name="TaxCode" style="width: 100%">
			<option></option>
			<% set rw = Server.CreateObject("ADODB.RecordSet")
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKGetTaxCodes" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@LanID") = Session("LanID")
			set rw = cmd.execute()
			do while not rw.eof %>
			<option <% If TaxCode = rw("Code") Then %>selected<% End If %> value="<%=myHTMLEncode(rw("Code"))%>"><%=myHTMLEncode(rw("Code"))%> 
			- <%=myHTMLEncode(rw("Name"))%></option>
			<% rw.movenext
			loop %>
			</select></td>
		</tr>
		<tr class="GeneralTbl">
			<td align="center" colspan="2">
			<input type="submit" value="<%=getAddCartGetTaxCodeLngStr("DtxtConfirm")%>" name="btnConfirm"> 
			- <input type="button" value="<%=getAddCartGetTaxCodeLngStr("DtxtCancel")%>" name="B2" onClick="javascript:window.close();"></td>
		</tr>
		</table>
</div>
<input type="hidden" name="Item" value="<%=myHTMLEncode(Request("Item"))%>">
<input type="hidden" name="T1" value="1">
<input type="hidden" name="redir" value="<% If Request("redir") = "searchCart" Then %>no<% Else %><%=Request("redir")%><% End If %>">
<input type="hidden" name="document" value="<%=Request("document")%>">
<input type="hidden" name="page" value="<%=Request("page")%>">
<input type="hidden" name="expItem" value="<%=Request("expItem")%>">
<input type="hidden" name="addPath" value="<%=Request("addPath")%>">
<input type="hidden" name="retVal" value="<%=Request("retVal")%>">
<input type="hidden" name="pop" value="Y">
</form>
<% Else %>
<script language="javascript">
<% If Request("expItem") <> "Y" Then %>
	opener.doMyLink('cart/addCartSubmitM.asp', 'item=<%=Request("Item")%>&T1=1&redir=<%=Request("redir")%>&TaxCode=<%=Request("TaxCode")%>&retVal=<%=Request("retVal")%>&document=<%=Request("document")%>&page=<%=Request("page")%>', '');
<% Else %>
	opener.doMyLink('cart/addCartSubmitExp.asp', 'item=<%=Request("Item")%>&redir=cart&TaxCode=<%=Request("TaxCode")%>&document=<%=Request("document")%>&page=<%=Request("page")%>', '');
<% End If %>
window.close();
</script>
<% End If %>
</body>
<% conn.close %>
</html>
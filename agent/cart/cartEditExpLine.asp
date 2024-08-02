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

<!--#include file="lang/cartEditExpLine.asp" -->
<!--#include file="../myHTMLEncode.asp"-->
<html <% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=getcartEditExpLineLngStr("LttlExpDet")%></title>
<link rel="stylesheet" type="text/css" href="../design/0/style/stylePopUp.css">
</head>

<body topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0">
<%
      set rs = Server.CreateObject("ADODB.recordset")
			sql = "select T0.ExpnsCode, IsNull(ExpnsName, '') ExpnsName, IsNull(Comments, '') Comments, TaxCode, VatGroup " & _
			"from R3_ObsCommon..DOC3 T0 " & _
			"inner join OEXD T1 on T1.ExpnsCode = T0.ExpnsCode " & _
			"where LogNum = " & Session("RetVal") & " and LineNum = " & Request("LineNum")
			set rs = conn.execute(sql)
			
If Request.Form.Count > 0 Then
	If Request("Comments") <> "" Then Comments = "N'" & saveHTMLDecode(Request("Comments"), False) & "'" Else Comments = "NULL"
	sql = "update R3_ObsCommon..DOC3 set Comments = " & Comments
	Select Case myApp.LawsSet 
		Case "PA", "AT", "AU", "BE", "CH", "CZ", "DE", "DK", "ES", "FI", "FR", "HU", "IT", "NL", "NO", "PL", "PT", "RU", "SE", "SK", "GB", "ZA"
			sql = sql & ", VatGroup = '" & saveHTMLDecode(Request("VatGroup"), False) & "' "
		Case "MX", "CL", "CR", "GT", "US", "CA"
			sql = sql & ", TaxCode = '" & saveHTMLDecode(Request("TaxCode"), False) & "' "
	End Select
	sql = sql & " where LogNum = " & Session("RetVal") & " and LineNum = " & Request("LineNum")
	conn.execute(sql) %>
<SCRIPT LANGUAGE="JavaScript">
opener.location.href = '../cart.asp?cmd=<%=Request("redir")%>';
window.close()
</script>
<% Else %>
<script language="javascript">
function chkMax(e, f, m)
{
	if(f.value.length == m && (e.keyCode != 8 && e.keyCode != 9 && e.keyCode != 35 && e.keyCode != 36 && e.keyCode != 37 
	&& e.keyCode != 38 && e.keyCode != 39 && e.keyCode != 40 && e.keyCode != 46 && e.keyCode != 16))return false; else return true;
}
</script>
<form method="POST" action="cartEditExpLine.asp" name="form1">
<div align="left">
	<table border="0" cellpadding="0" width="399" id="table1">
		<tr class="GeneralTbl">
			<td  class="GeneralTblBold2" width="393" align="center" colspan="2"><%=rs("ExpnsCode")%> 
			- <%=rs("ExpnsName")%>&nbsp;</td>
		</tr>
		<% Select Case myApp.LawsSet
			Case "PA", "AT", "AU", "BE", "CH", "CZ", "DE", "DK", "ES", "FI", "FR", "HU", "IT", "NL", "NO", "PL", "PT", "RU", "SE", "SK", "GB", "ZA" %>
		<tr class="GeneralTbl">
			<td  class="GeneralTblBold2" width="92" align="center">
			<%=getcartEditExpLineLngStr("LtxtVatGrp")%>:</td>
			<td width="301">
			<% sql = "select Code, IsNull(Name, '') Name from OVTG where Category = 'O'"
        	set rw = conn.execute(sql) %>
			<select size="1" name="VATGroup" style="width: 100%">
			<% do while not rw.eof %>
			<option value="<%=RW(0)%>" <% If Rw(0) = rs("VATGroup") Then %>selected<%end if %>><%=myHTMLEncode(RW(1))%></option>
			<% rw.movenext
			loop %>
			</select></td>
		</tr>
		<% Case "MX", "CL", "CR", "GT", "US", "CA" %>
		<tr class="GeneralTbl">
			<td  class="GeneralTblBold2" width="92" align="center">
			<%=getcartEditExpLineLngStr("LtxtTaxCode")%>:</td>
			<td width="301"><select size="1" name="TaxCode" style="width: 100%">
			<% 
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKGetTaxCodes" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@LanID") = Session("LanID")
			set rw = cmd.execute()
			do while not rw.eof %>
			<option <% If rs("TaxCode") = rw("Code") Then %>selected<% End If %> value="<%=myHTMLEncode(rw("Code"))%>"><%=myHTMLEncode(rw("Code"))%> 
			- <%=myHTMLEncode(rw("Name"))%></option>
			<% rw.movenext
			loop %>
			</select></td>
		</tr>
      <% End Select %>
		<tr class="GeneralTbl">
			<td  class="GeneralTblBold2" width="92" align="center"><%=getcartEditExpLineLngStr("DtxtNote")%>:</td>
			<td width="301"><textarea rows="5" name="Comments" cols="47" onkeydown="return chkMax(event, this, 254);"><% If Not IsNull(rs("Comments")) Then %><%=myHTMLEncode(rs("Comments"))%><% End If %></textarea></td>
		</tr>
		<tr class="GeneralTbl">
			<td colspan="2">
			<table cellpadding="0" cellspacing="0" border="0" width="100%">
				<tr>
					<td><input type="submit" value="<%=getcartEditExpLineLngStr("DtxtConfirm")%>" name="B1"></td>
					<td align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><input type="submit" value="<%=getcartEditExpLineLngStr("DtxtCancel")%>" name="B2"></td>
				</tr>
			</table>
			</td>
		</tr>
	</table>
</div>
<input type="hidden" name="LineNum" value="<%=Request("LineNum")%>">
<input type="hidden" name="redir" value="<%=Request("redir")%>">
<input type="hidden" name="AddPath" value="../">
<input type="hidden" name="pop" value="Y">
</form>

<% End If %>
</body>
<% conn.close %>
</html>
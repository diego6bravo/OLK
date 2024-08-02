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
<!--#include file="lang/clientSearch.asp" -->
<%
sourceCount = 0
targetCount = 0 %>
<!--#include file="../loadAlterNames.asp" -->
<!--#include file="../myHTMLEncode.asp"-->
<%       
           	set rs = Server.CreateObject("ADODB.recordset")
	  		If Request("B4") = "" Then
      			If Request("B1") <> "" Then
      				sql = ""
	      			ArrVal = Split(Request.Form("SourceCardCode"),", ")
	      			For i = 0 to UBound(ArrVal)
		      			sql = sql & "if not exists(select 'A' from OLKMsgTemp where SlpCode = " & Session("vendid") & " and CardCode = N'" & saveHTMLDecode(ArrVal(i), False) & "') begin insert olkmsgtemp(SlpCode, CardCode) Values(" & Session("vendid") & ", N'" & saveHTMLDecode(ArrVal(i), False) & "') end " & VbCrLf
	      			next
      			conn.execute(sql)
      		ElseIf Request("B2") <> "" Then
      			sql = ""
	      		ArrVal = Split(Request.Form("CardCode"),", ")
	      		For i = 0 to UBound(ArrVal)
		      		sql = sql & "delete olkmsgtemp where SlpCode = " & Session("vendid") & " and CardCode = N'" & saveHTMLDecode(ArrVal(i), False) & "' " & VbCrLf
	      		next
      			conn.execute(sql)
      		End If
      
      		If Request("Group") <> "" Then Group = " and GroupCode = " & Request("Group")
      		If Request("Country") <> "" Then Country = " and Country = '" & Request("Country") & "'"
      		sql3 = "select T0.CardCode, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRD', 'CardName', T0.CardCode, T0.CardName) CardName from ocrd T0 " & _
      				"inner join OLKCLientsAccess T1 on T1.CardCode = T0.CardCode " & _
      				"where Status = 'A' and (T0.CardCode like N'%" & saveHTMLDecode(Request("String"), False) & "%' or CardName like N'%" & saveHTMLDecode(Request("String"), False) & "%') " & _
	        		Group & Country & " and T0.CardCode Not In (Select CardCode from olkMsgTemp where SlpCode = " & Session("vendid") & ")"
	        
	  		sql4 = "select T0.CardCode, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRD', 'CardName', T0.CardCode, T0.CardName) CardName from ocrd T0 inner join OLKMsgTemp T1 on T1.CardCode = T0.CardCode and T1.SlpCode = " & Session("vendid")
	  
           %>
<html <% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=Replace(getclientSearchLngStr("LttlCSearch"), "{0}", Server.HTMLEncode(LCase(txtClients)))%></title>
<link rel="stylesheet" type="text/css" href="../design/0/style/stylePopUp.css">
<script language="javascript" src="../general.js"></script>
</head>
<body topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" bgcolor="#FFFFFF">
<script language="javascript">
function confirmSubmit(Button) {
if (Button.name == 'B1' && document.frmCardCode.sourceCardCode.value == '') 
{
	alert('<%=getclientSearchLngStr("LtxtValSelCard")%>'.replace('{0}', '<%=txtClient%>'));
	return false; 
}

if (Button.name == 'B2' && document.frmCardCode.CardCode.value == '') 
{
	alert('<%=getclientSearchLngStr("LtxtValRemCard")%>'.replace('{0}', '<%=txtClient%>'))
	return false 
}
return true
}

function selectAll(ListBox, ListCount, B)
{
	for (var i = 0;i<ListCount;i++)
		ListBox.options[i].selected = true;
		
	B.click();
}
</script>
<table border="0" cellpadding="0" width="500" id="table1">
	<tr class="GeneralTlt">
		<td>
		<p align="center"><% If 1 = 2 Then %>Clientes<% Else %><%=txtClients%><% End If %></td>
	</tr>
	<form method="POST" action="clientSearch.asp" webbot-action="--WEBBOT-SELF--">
	<input type="hidden" name="search" value="Y">
	<input type="hidden" name="AddPath" value="../">
	<input type="hidden" name="pop" value="Y">
	</form>
	<form method="POST" action="clientSearch.asp" name="frmCardCode" webbot-action="--WEBBOT-SELF--">
	<tr class="GeneralTbl">
		<td>
		<table border="0" cellpadding="0" width="100%" id="table5">
			<tr>
				<td class="GeneralTblBold2"><%=getclientSearchLngStr("DtxtSearch")%></td>
				<td class="GeneralTbl">
				<input type="text" name="string" size="34" value="<%=Request("string")%>"></td>
				<td class="GeneralTbl">&nbsp;</td>
			</tr>
			<tr>
				<td class="GeneralTblBold2"><%=getclientSearchLngStr("DtxtGroup")%></td>
				<td class="GeneralTbl"><select size="1" name="Group">
				<option></option>
				<% 				
				set cmd = Server.CreateObject("ADODB.Command")
				cmd.ActiveConnection = connCommon
				cmd.CommandType = &H0004
				cmd.CommandText = "DBOLKGetCrdGroups" & Session("ID")
				cmd.Parameters.Refresh()
				cmd("@LanID") = Session("LanID")
				cmd("@CardType") = "C"
				set rs = cmd.execute()
				do while not rs.eof %>
				<option <% If CStr(Request("Group")) = CStr(rs("GroupCode")) then %>selected<%end if %> value="<%=rs("GroupCode")%>"><%=myHTMLEncode(rs("GroupName"))%></option>
				<% rs.movenext
				loop %>
				</select></td>
				<td class="GeneralTbl">&nbsp;</td>
			</tr>
			<tr>
				<td class="GeneralTblBold2"><%=getclientSearchLngStr("DtxtCountry")%></td>
				<td class="GeneralTbl"><select size="1" name="Country">
				<option></option>
				<% set cmd = Server.CreateObject("ADODB.Command")
				cmd.ActiveConnection = connCommon
				cmd.CommandType = &H0004
				cmd.CommandText = "DBOLKGetCountries" & Session("ID")
				cmd.Parameters.Refresh()
				cmd("@LanID") = Session("LanID")
				set rs = cmd.execute()
				do while not rs.eof %>
				<option <% If CStr(Request("Country")) = CStr(rs("code")) then %>selected<%end if %> value="<%=rs("code")%>"><%=myHTMLEncode(rs("name"))%></option>
				<% rs.movenext
				loop %>
				</select></td>
				<td class="GeneralTbl">
				<p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
				<input type="submit" value="<%=getclientSearchLngStr("DbtnSearch")%>" name="B3"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table3">
			<tr class="GeneralTblBold2">
				<td width="52%" align="center"><%=getclientSearchLngStr("DtxtSearch")%></td>
				<td align="center" width="47%"><%=getclientSearchLngStr("LtxtSendTo")%>:</td>
			</tr>
			<tr class="GeneralTbl">
				<td colspan="2">
				<table border="0" cellpadding="0" width="100%" cellspacing="0" id="table4">
					<tr>
						<td width="309">
						<p align="center"><select style="width: 200" size="10" name="sourceCardCode" multiple ondblclick="vbscript:document.frmCardCode.B1.click()">
						<% If Request.Form("search") = "Y" then
						rs.close
						rs.open sql3, conn, 3, 1
						sourceCount = rs.recordcount
						do while not rs.eof %>
						<option value="<%=myHTMLEncode(rs("CardCode"))%>"><%=myHTMLEncode(rs("CardCode"))%> - <%=myHTMLEncode(rs("CardName"))%></option>
						<% rs.movenext
						loop
						end if %>
						</select></td>
						<td width="110">
						<p align="center">
						<% If sourceCount > 0 Then %>
						<input type="button" value="-&gt;&gt;" name="B5" onclick="javascript:selectAll(document.frmCardCode.sourceCardCode, <%=sourceCount-1%>, document.frmCardCode.B1)"><br>
						<br>
						<input type="submit" value="--&gt;" name="B1" onclick="return confirmSubmit(this)"><br>
						<br><% End If %>
						<% rs.close
						rs.open sql4, conn, 3, 1 
						targetCount = rs.recordcount 
						if targetCount > 0 then %>
						<input type="submit" value="&lt;--" name="B2" onclick="return confirmSubmit(this)"><br>
						<br>
						<input type="button" value="&lt;&lt;-" name="B6" onclick="javascript:selectAll(document.frmCardCode.CardCode, <%=targetCount-1%>, document.frmCardCode.B2)">
						<% End If %></td>
						<td>
						<p align="center"><select style="width: 200" size="10" name="CardCode" multiple ondblclick="vbscript:document.frmCardCode.B2.click()">
						<% do while not rs.eof %>
						<option value="<%=myHTMLEncode(rs("CardCode"))%>"><%=myHTMLEncode(rs("CardCode"))%> - <%=myHTMLEncode(rs("CardName"))%></option>
						<% rs.movenext
						loop %>
						</select></td>
					</tr>
				</table>
				</td>
			</tr>
		</table>
		</td>
	</tr>
	<tr class="GeneralTbl">
		<td>
		<p align="center"><input type="submit" value="<%=getclientSearchLngStr("DtxtAccept")%>" name="B4"></td>
	</tr>
		<input type="hidden" name="search" value="Y">
		<input type="hidden" name="AddPath" value="../">
		<input type="hidden" name="pop" value="Y">
		</form>
</table>
</body>

</html>
<% Else 
sql = "select CardCode from olkMsgTemp where slpcode = " & Session("vendid")
set rs = Server.CreateObject("ADODB.RecordSet")
rs.open sql, conn, 3, 1
do while not rs.eof
varx = varx & rs("CardCode")
If rs.bookmark <> rs.recordcount then varx = varx & ", "
rs.movenext
loop %>
<script language="javascript" src="../general.js"></script>
<script language="javascript">
opener.clientsTo('<%=Replace(myHTMLEncode(varx), "'", "\'")%>');
window.close();
</script>
<% End If
conn.close
set rs = nothing %>
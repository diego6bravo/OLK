<!--#include file="top.asp" -->
<!-- #INCLUDE file="FCKeditor/fckeditor.asp" -->
<!--#include file="lang/adminObjPrint.asp" -->
<!--#include file="adminTradSubmit.asp"-->
<!--#include file="accountControl.asp"-->  
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style type="text/css">
.style1 {
	background-color: #E2F3FC;
	font-family: Verdana;
	font-size: xx-small;
	color: #31659C;
}
.style2 {
	background-color: #F7FBFF;
}
.style3 {
	font-weight: bold;
	background-color: #E1F3FD;
}
</style>
</head>
<script language="javascript" src="js_up_down.js"></script>
<%
conn.execute("use [" & Session("olkdb") & "]")

obj = -1

If Request("object") <> "" Then obj = Request("object")
%>

<form method="POST" action="adminSubmit.asp" name="frmPrint" onsubmit="javascript:return valFrm()">
<input type="hidden" name="submitCmd" value="adminPrint">
<input type="hidden" name="cmd" value="U">
<table border="0" cellpadding="0" width="100%">
	<tr>
		<td bgcolor="#E1F3FD"><b><font color="#FFFFFF">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><%=getadminObjPrintLngStr("LttlObjPrint")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif">
		<font face="Verdana" size="1" color="#4783C5"><%=getadminObjPrintLngStr("LttlObjPrintNote")%></font></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr>
				<td class="style1">
				<select size="1" name="ObjectCode" onchange="javascript:window.location.href='adminObjPrint.asp?object='+this.value" class="input">
				<option value=""><%=getadminObjPrintLngStr("LtxtSelObj")%></option>
				<optgroup label="<%=getadminObjPrintLngStr("LtxtGeneral")%>">
					<option <% If obj = 33 Then %>selected<% End If %> value="33">
					<%=getadminObjPrintLngStr("DtxtActivity")%></option>
					<option <% If obj = 4 Then %>selected<% End If %> value="4">
					<%=getadminObjPrintLngStr("DtxtItem")%></option>
					<option <% If obj = 2 Then %>selected<% End If %> value="2">
					<%=getadminObjPrintLngStr("DtxtBPS")%></option>					
				</optgroup>
				<optgroup label="<%=getadminObjPrintLngStr("LtxtPurchase")%>">
					<option <% If obj = 20 Then %>selected<% End If %> value="20">
					<%=getadminObjPrintLngStr("LtxtGoodsIssue")%></option>
					<option <% If obj = 22 Then %>selected<% End If %> value="22">
					<%=getadminObjPrintLngStr("DtxtPurOrder")%></option>
				</optgroup>
				<optgroup label="<%=getadminObjPrintLngStr("LtxtSale")%>">
					<option <% If obj = 23 Then %>selected<% End If %> value="23">
					<%=getadminObjPrintLngStr("DtxtQuote")%></option>
					<option <% If obj = 17 Then %>selected<% End If %> value="17">
					<%=getadminObjPrintLngStr("DtxtSalesOrder")%></option>
					<option <% If obj = 15 Then %>selected<% End If %> value="15">
					<%=getadminObjPrintLngStr("DtxtDelivery")%></option>
					<option <% If obj = 203 Then %>selected<% End If %> value="203">
					<%=getadminObjPrintLngStr("DtxtARDownPayReq")%></option>
					<option <% If obj = 204 Then %>selected<% End If %> value="204">
					<%=getadminObjPrintLngStr("DtxtARDownPayInv")%></option>
					<option <% If obj = 13 Then %>selected<% End If %> value="13">
					<%=getadminObjPrintLngStr("DtxtInvoice")%></option>
					<option <% If obj = -13 Then %>selected<% End If %> value="-13">
					<%=getadminObjPrintLngStr("DtxtInvoice")%> (<%=getadminObjPrintLngStr("DtxtReservada")%>)</option>
					<option <% If obj = 48 Then %>selected<% End If %> value="48">
					<%=getadminObjPrintLngStr("DtxtInvoice")%>/<%=getadminObjPrintLngStr("DtxtReceipt")%></option>
				</optgroup>
				<optgroup label="<%=getadminObjPrintLngStr("LtxtBanks")%>">
					<option <% If obj = 24 Then %>selected<% End If %> value="24">
					<%=getadminObjPrintLngStr("DtxtReceipt")%></option>
				</optgroup>
				</select></td>
			</tr>
			<% If obj <> -1 Then
			
			set cmdSec = Server.CreateObject("ADODB.Command")
			cmdSec.ActiveConnection = connCommon
			cmdSec.CommandType = &H0004
			cmdSec.CommandText = "DBOLKGetObjectPrintAvlSec" & Session("ID")
			cmdSec.Parameters.Refresh()
			cmdSec("@ObjectCode") = obj
			set rd = Server.CreateObject("ADODB.RecordSet")
			
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKGetObjectPrintList" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@ObjectCode") = obj
			set rs = Server.CreateObject("ADODB.RecordSet")
			rs.open cmd, , 3, 1
			For i = 0 to 2
			
			UserType = ""
			Select Case i
				Case 0
					UserType = "A"
				Case 1
					UserType = "C"
				Case 2
					UserType = "P"
			End Select
			
			
			rs.Filter = "UserType = '" & UserType & "'" %>
			<tr>
				<td class="style1">
				<font color="#4783C5" face="Verdana" size="1"><% Select Case i
				Case 0 %><%=getadminObjPrintLngStr("DtxtAgent")%><%
				Case 1 %><%=getadminObjPrintLngStr("DtxtClient")%><%
				Case 2 %><%=getadminObjPrintLngStr("DtxtPocket")%><% 
				End Select %></font></td>
			</tr>
			<tr>
				<td>
					<table border="0" cellpadding="0" width="100%">
						<tr>
							<td align="center" class="style3" style="width: 240">
							<font face="Verdana" size="1" color="#31659C"><%=getadminObjPrintLngStr("DtxtName")%></font></td>
							<td align="center" class="style3">
							<font face="Verdana" size="1" color="#31659C"><%=getadminObjPrintLngStr("LLinkData")%></font></td>
							<td align="center" width="70" class="style3">
							<font face="Verdana" size="1" color="#31659C"><%=getadminObjPrintLngStr("DtxtOrder")%></font></td>
							<td align="center" width="60" class="style3">
							<font face="Verdana" size="1" color="#31659C"><%=getadminObjPrintLngStr("DtxtActive")%></font></td>
							<td width="20" class="style2"><font size="1">&nbsp;</font></td>
						</tr>
						<% 	do while not rs.eof
							SecID = CInt(rs("SecID"))
							myID = Replace(SecID, "-", "_")
							If SecID >= 0 Then 
								SecName = rs("SecName")
							Else 
								If Not IsNull(rs("LinkName")) Then
									SecName = rs("LinkName")
								Else
									SecName = getadminObjPrintLngStr("LtxtExternalLink")
								End If
							End If %>
							<input type="hidden" name="SecID" value="<%=myID%>">
						<tr>
							<td bgcolor="#F3FBFE" style="width: 240"><% 
							If SecID >= 0 Then
								%><font face="Verdana" color="#4783C5" size="1"><%=SecName%></font><%
							Else %>
							<table cellpadding="0" cellspacing="0" border="0" style="width: 100%;">
								<tr>
									<td><input class="input" type="text" name="LinkName<%=myID%>" size="44" style="width: 100%;" maxlength="50" value="<%=Server.HTMLEncode(SecName)%>"></td>
									<td style="width: 16px;"><a href="javascript:doFldTrad('ObjectPrint', 'ObjectCode, SecID', '<%=obj%>,<%=SecID%>', 'AlterLinkName', 'T', null);"><img src="images/trad.gif" alt="<%=getadminObjPrintLngStr("DtxtTranslate")%>" border="0"></a></td>
								</tr>
							</table><% End If %></td>
							<td bgcolor="#F3FBFE"><input class="input" style="width: 100%;" type="text" id="LinkData" name="LinkData<%=myID%>" size="54" value="<%=Server.HTMLEncode(rs("LinkData"))%>"></td>
							<td width="70" bgcolor="#F3FBFE">
							<font face="Verdana" size="1">
							<table cellpadding="0" cellspacing="0" border="0">
								<tr>
									<td><input id="SecOrder<%=myID%>" name="Order<%=myID%>" class="input" value="<%=rs("Ordr")%>" size="7" onfocus="this.select();" onmouseup="event.preventDefault()" onkeydown="return chkMax(event, this, 6);"></td>
									<td valign="middle">
									<table cellpadding="0" cellspacing="0" border="0">
										<tr>
											<td><img src="images/img_nud_up.gif" id="btnSecOrder<%=myID%>Up"></td>
										</tr>
										<tr>
											<td><img src="images/spacer.gif"></td>
										</tr>
										<tr>
											<td><img src="images/img_nud_down.gif" id="btnSecOrder<%=myID%>Down"></td>
										</tr>
									</table>
									</td>
								</tr>
							</table>
							<script language="javascript">NumUDAttach('frmSec', 'SecOrder<%=myID%>', 'btnSecOrder<%=myID%>Up', 'btnSecOrder<%=myID%>Down');</script>
							</td>
							<td bgcolor="#F3FBFE">
							<p align="center">
							<input <% If rs("Active") = "Y" Then %>checked <% End If %> type="checkbox" name="Active<%=myID%>" value="Y" class="noborder"></td>
							<td width="20" bgcolor="#F3FBFE"><a href="javascript:delPrint('<%=Replace(SecName, "'", "\'")%>', <%=rs("SecID")%>);"><img border="0" src="images/remove.gif" width="16" height="16"></a></td>
						</tr>
						<% 
						rs.movenext
						loop %>
					</table>
				</td>
			</tr>
			<tr>
				<td>
					<table border="0" cellpadding="0" width="100%">
						<tr>
							<td class="style3" style="width: 240"><font face="Verdana" size="1" color="#31659C"><%=getadminObjPrintLngStr("DtxtAdd")%></font></td>
							<td bgcolor="#F3FBFE">
							<select name="cmbAddSec" onchange="goAdd(this.value, '<%=UserType%>');">
							<option value=""></option>
							<option value="-1"> - <%=getadminObjPrintLngStr("LtxtExternalLink")%> -</option>
							<% 
							cmdSec("@UserType") = UserType
							set rd = cmdSec.execute()
							do while not rd.eof %>
							<option value="<%=rd("SecID")%>"><%=rd("SecName")%></option>
							<% rd.movenext
							loop %>
							</select></td>
						</tr>
					</table>
				</td>
			</tr>
			<% Next %>
			<% End If %>
		</table>
		</td>
	</tr>
	<% If obj <> -1 Then %>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr>
				<td width="77">
				<input type="submit" value="<%=getadminObjPrintLngStr("DtxtSave")%>" name="B1" class="OlkBtn"></td>
				<td><hr color="#0D85C6" size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
	<% End If %>
</table>
<script language="javascript">
function delPrint(secName, secID)
{
	if(confirm('<%=getadminObjPrintLngStr("LtxtConfRemPrint")%>'.replace('{0}', secName)))window.location.href='adminSubmit.asp?submitCmd=adminPrint&cmd=D&delID=' + secID +'&ObjectCode=<%=obj%>';
}
function goAdd(secID, userType)
{
	if (secID != "")
		window.location.href='adminSubmit.asp?submitCmd=adminPrint&cmd=A&secID=' + secID +'&ObjectCode=<%=obj%>&userType=' + userType;
}

function valFrm()
{
	return true;
}
</script>
<% If Session("style") = "nc" Then %>
<br>
	<input name="adminSec" type="hidden" value="adminPrint">
<% End If %>
<!--#include file="bottom.asp" -->
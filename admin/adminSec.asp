<!--#include file="top.asp" -->
<!--#include file="lang/adminSec.asp" -->
<!--#include file="adminTradSubmit.asp"-->
<% conn.execute("use [" & Session("olkdb") & "]") %>
<script language="javascript">
var txtValSecNam 			= '<%=getadminSecLngStr("LtxtValSecNam")%>';
var txtValActiveDep 		= '<%=getadminSecLngStr("LtxtValActiveDep")%>';
var txtValActiveDepEnd 		= '<%=getadminSecLngStr("LtxtValActiveDepEnd")%>';
var txtValDeactiveSec 		= '<%=getadminSecLngStr("LtxtValDeactiveSec")%>';
var txtValDeactiveSecNow 	= '<%=getadminSecLngStr("LtxtValDeactiveSecNow")%>';
</script>
<script language="javascript" src="adminSec.js"></script>
<script language="javascript" src="js_up_down.js"></script>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style type="text/css">
.style2 {
	background-color: #E1F3FD;
}
.style3 {
	font-weight: bold;
	background-color: #E1F3FD;
}
</style>
</head>

<table border="0" cellpadding="0" width="100%" id="table3">
	<form method="POST" action="adminSubmit.asp" name="frmSec" onsubmit="return valSecFrm()">
	<tr>
		<td bgcolor="#E1F3FD"><b><font face="Verdana" size="2">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><% 
		Select Case Request("UType") 
			Case "C" %><%=getadminSecLngStr("LttlDefOlkSec")%>
		<% 	Case "P" %><%=getadminSecLngStr("LttlDefOlkForm")%>
		<%	Case "A" %><%=getadminSecLngStr("LttlDefOlkAgent")%>
		<% End Select %></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1">
		<font color="#4783C5"><% Select Case Request("UType") 
			Case "C" %><%=getadminSecLngStr("LttlDefOlkSecNote")%>
		<% 	Case "P" %><%=getadminSecLngStr("LttlDefOlkFormNote")%>
		<%	Case "A" %><%=getadminSecLngStr("LttlDefOlkAgentNote")%>
		<% End Select %></font></font></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table11">
			<tr>
				<td width="20" class="style2"><font size="1">&nbsp;</font></td>
				<td align="center" width="70" class="style3">
				<font face="Verdana" size="1" color="#31659C"><%=getadminSecLngStr("DtxtType")%></font></td>
				<td align="center" class="style3">
				<font face="Verdana" size="1" color="#31659C"><%=getadminSecLngStr("DtxtName")%></font></td>
				<td align="center" width="70" class="style3">
				<font face="Verdana" size="1" color="#31659C"><%=getadminSecLngStr("DtxtOrder")%></font></td>
				<% If Request("UType") <> "P" Then %>
				<td align="center" width="60" class="style3">
				<font face="Verdana" size="1" color="#31659C"><%=getadminSecLngStr("LtxtSesion")%></font></td><% End If %>
				<td align="center" width="90" class="style3">
				<font face="Verdana" size="1" color="#31659C"><%=getadminSecLngStr("LtxtHideMenu")%></font></td>
				<% If Request("UType") = "C" Then %>
				<td align="center" width="90" class="style3">
				<font face="Verdana" size="1" color="#31659C"><%=getadminSecLngStr("LtxtHideSecMenu")%></font></td><% End If %>
				<td align="center" width="60" class="style3">
				<font face="Verdana" size="1" color="#31659C"><%=getadminSecLngStr("LtxtActive")%></font></td>
				<td width="20" class="style2"><font size="1">&nbsp;</font></td>
			</tr>
			<% 	sql = "select T0.SecType, T0.SecID, Case T0.SecType When 'S' Then T1.SecName collate database_default Else T0.SecName End SecName,  " & _
				"T0.SecOrder, T0.ReqLogin, T0.Status, T0.HideMainMenu, T0.HideSecondMenu, T0.Type " & _
				"from OLKSections T0 " & _
				"left outer join OLKCommon..OLKSectionsDesc T1 on T1.SecID = T0.SecID and T0.SecType = 'S' and T1.LanID = " & Session("LanID") & " " & _
				"where T0.Status <> 'D' and T0.UserType = '" & Request("UType") & "' " & _
				"order by T0.SecOrder asc "
				rs.open sql, conn, 3, 1
				NewOrder = 0
				do while not rs.eof
				
				myID = rs("SecType") & rs("SecID") %>
				<input type="hidden" id="Type" name="Type<%=myID%>" value="<%=rs("Type")%>">
			<tr>
				<td width="20" bgcolor="#F3FBFE"><% If rs("SecType") = "U" or rs("SecType") = "S" and (rs("SecID") = -6 or rs("SecID") = -7 or rs("SecID") = -3) Then
				If rs("SecType") = "U" Then 
					lnk = "adminSecEdit.asp?SecID=" & rs("SecID") & "&rCount=" & rs.recordcount & "&UType=" & Request("UType")
				ElseIf rs("SecType") = "S" and rs("SecID") = -7 Then 
					lnk = "adminSecIndex.asp"
				ElseIf rs("SecType") = "S" and rs("SecID") = -6 Then 
					lnk = "adminMyData.asp"
				ElseIf rs("SecType") = "S" and rs("SecID") = -3 Then
					lnk = "adminCartMore.asp"
				End If %>
				<a href="<%=lnk%>">
				<img border="0" src="images/<%=Session("rtl")%>flechaselec.gif" width="15" height="13"></a><% End If %></td>
				<td width="70" bgcolor="#F3FBFE" align="center">
				<font color="#4783C5" face="Verdana" size="1"><% 
				Select Case rs("SecType")
					Case "S"
						Response.Write getadminSecLngStr("DtxtSystem")
					Case "U"
						Select Case rs("Type")
							Case "N"
								Response.Write getadminSecLngStr("DtxtContent")
							Case "R"
								Response.Write getadminSecLngStr("DtxtReport")
							Case "L"
								Response.Write getadminSecLngStr("DtxtLink")
						End Select
					End Select %></font></td>
				<td bgcolor="#F3FBFE">
				<table cellpadding="0" cellspacing="0" border="0" width="100%">
					<tr>
						<td><input class="input" <% If rs("SecType") = "S" Then %>disabled<% End If %> style="width: 100%; <% If rs("SecType") = "S" Then %>background-color: #D9F0FD;<% End If %> " type="text" id="SecName" name="SecName<%=myID%>" size="54" value="<%=Server.HTMLEncode(rs("SecName"))%>" onkeydown="return chkMax(event, this, 100);"></td>
						<td style="width: 16px"><% If rs("SecType") = "U" Then %><a href="javascript:doFldTrad('Sections', 'SecType,SecID', '<%=rs("SecType")%>,<%=rs("SecID")%>', 'AlterSecName', 'T', null);"><img src="images/trad.gif" alt="<%=getadminSecLngStr("DtxtTranslate")%>" border="0"></a><% Else %>&nbsp;<% End If %></td>
					</tr>
				</table>
				</td>
				<td width="70" bgcolor="#F3FBFE">
				<font face="Verdana" size="1">
				<% If rs("SecOrder") >= 0 Then %>
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td><input id="SecOrder<%=myID%>" name="SecOrder<%=myID%>" class="input" value="<%=rs("SecOrder")%>" size="7" onfocus="this.select();" onmouseup="event.preventDefault()" onkeydown="return chkMax(event, this, 6);"></td>
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
				<% Else %>
				<input type="hidden" name="SecOrder<%=myID%>" value="<%=rs("SecOrder")%>">
				<% End If %>
				</td>
				<% If Request("UType") <> "P" Then %>
				<td bgcolor="#F3FBFE">
				<p align="center">
				<input <% If (rs("SecType") = "S" and rs("SecID") <> 0) or rs("Type") = "R" and Request("UType") = "C" Then %>disabled <% End If %><% If rs("ReqLogin") = "Y" Then %>checked<% End If %> type="checkbox" name="ReqLogin<%=myID%>" value="Y" class="noborder"></td><% End If %>
				<td bgcolor="#F3FBFE">
				<p align="center">
				<input <% If rs("SecType") = "S" and rs("SecID") <> 0 and rs("SecID") <> 3 Then %>disabled <% End If %> <% If rs("HideMainMenu") = "Y" Then %>checked<% End If %> type="checkbox" name="HideMainMenu<%=myID%>" value="Y" class="noborder"></td>
				<% If Request("UType") = "C" Then %>
				<td bgcolor="#F3FBFE">
				<p align="center">
				<input <% If rs("SecType") = "S" and rs("SecID") <> 3 Then %>disabled <% End If %> <% If rs("HideSecondMenu") = "Y" Then %>checked<% End If %> type="checkbox" name="HideSecondMenu<%=myID%>" value="Y" class="noborder"></td><% End If %>
				<td bgcolor="#F3FBFE">
				<p align="center">
				<input <% If rs("Status") = "A" Then %>checked <% End If %> type="checkbox" name="Status<%=myID%>" value="Y" <% If rs("SecType") = "S" Then %>onclick="javascript:chkSec(this);"<% End If %> class="noborder"></td>
				<td width="20" bgcolor="#F3FBFE"><% If rs("SecType") = "U" Then %><a href="javascript:delSec('<%=Replace(rs("SecName"), "'", "\'")%>', <%=rs("SecID")%>);"><img border="0" src="images/remove.gif" width="16" height="16"></a><% End If %></td>
			</tr>
			<% 
			If rs.bookmark = rs.recordcount Then NewOrder = rs("SecOrder")+1
			rs.movenext
			loop %>
		</table>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table5">
			<tr>
				<td width="77">
				<input type="submit" value="<%=getadminSecLngStr("DtxtSave")%>" name="B1" class="OlkBtn"></td>
				<td width="77">
				<input type="button" value="<%=getadminSecLngStr("DtxtNew")%>" name="btnAdd" class="OlkBtn" onclick="javascript:window.location.href='adminSecEdit.asp?NewOrder=<%=NewOrder%>&rCount=<%=rs.recordcount%>&UType=<%=Request("UType")%>';"></td>
				<td><hr color="#0D85C6" size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
		<input type="hidden" name="uCmd" value="update">
		<input type="hidden" name="UType" value="<%=Request("UType")%>">
	<input type="hidden" name="submitCmd" value="adminSec">
	</form>
</table>
<script language="javascript">
function delSec(SecName, SecID)
{
	if(confirm('<%=getadminSecLngStr("LtxtConfDelSec")%>'.replace('{0}', SecName)))window.location.href='adminSubmit.asp?submitCmd=adminSec&uCmd=del&SecID=' + SecID +'&UType=<%=Request("UType")%>';
}
</script>
<!--#include file="bottom.asp" -->
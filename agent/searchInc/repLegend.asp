<% addLngPathStr = "searchInc/" %>
<!--#include file="lang/repLegend.asp" -->
<% 
set rLegend = Server.CreateObject("ADODB.RecordSet")
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetRepLegend" & Session("ID")
cmd.Parameters.Refresh()
cmd("@LanID") = Session("LanID")
cmd("@rsIndex") = CInt(Request("rsIndex"))
set rLegend = cmd.execute()
If Not rLegend.Eof Then
doRepLegend = True %>
<script language="javascript">
function showLegend(td)
{
	myTbl = document.getElementById('tblLegend');
	myTbl.style.display='';
	<% If userType = "V" Then %>
	myTbl.style.left = <% If Session("rtl") = "" Then %>0<% Else %>-myTbl.offsetWidth+td.offsetWidth+9<% End If %>;
	<% ElseIf userType = "C" Then %>
	myTbl.style.left = GetLeftPos(td)<% If Session("rtl") = "" Then %>-(myTbl.offsetWidth-td.offsetWidth)<% End If %>;
	<% End If %>
}
function hideLegend()
{
	document.getElementById('tblLegend').style.display='none';
}
</script>
<table border="0" width="100%" id="btnShowLegend" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; cursor: help">
	<tr onmouseover="showLegend(this);" onmouseout="hideLegend();">
		<td class="TblTlt" style="border: 1px solid #FFFFFF">
		<p align="center"><%=getrepLegendLngStr("LtxtLegend")%></td>
	</tr>
</table>
<table cellpadding="0" cellspacing="2" border="0" bgcolor="#FFFFFF" id="tblLegend" style="background-color: #FFFFFF; position: absolute; display: none; z-index: 1;">
	<% do while not rLegend.eof %>
	<tr class="TablasNoticias">
		<td class="rs_<%=rLegend("ColorID")%>_<%=rLegend("LineID")%>"><% If rLegend("FontBlink") = "Y" Then %><blink><% End If %><%=Replace(Server.HTMLEncode(rLegend("Alias")), " ", "&nbsp;")%><% If rLegend("FontBlink") = "Y" Then %></blink><% End If %></td>
	</tr>
	<% rLegend.movenext
	loop %>
</table>
<% Else
	doRepLegend = False
End If %>
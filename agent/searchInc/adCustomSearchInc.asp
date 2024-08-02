<% addLngPathStr = "searchInc/" %>
<!--#include file="lang/adCustomSearchInc.asp" -->
<%
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetCustomSearch" & Session("ID")
cmd.Parameters.Refresh()
cmd("@LanID") = Session("LanID")
cmd("@ObjID") = ObjID
set rs = Server.CreateObject("ADODB.RecordSet")
rs.open cmd, , 3, 1
If rs.recordcount = 1 Then %>
<input type="button" value="<%=getadCustomSearchIncLngStr("LtxtAdvanced")%>" name="B3" style="color: #FFFFFF; font-family: Verdana; font-size: 7pt; border: 1px solid #FFFFFF; background-color: #0065CE; width:76px;" onclick="goAdSearch(<%=rs("ID")%>, <%=ObjID%>);">
<% ElseIf rs.recordcount > 1 Then
If Session("RetVal") <> "" Then pageID = 1 Else pageID = 0 %>
<table cellpadding="0" cellspacing="2" border="0" style=" border: 1px solid #FFFFFF; background-color: #0065CE; width:105px; cursor: pointer;">
	<tr onmouseover="showAdSearch(this, <%=pageID%>);" onmouseout="hideAdSearch();">
		<td style="color: #FFFFFF; font-family: Verdana; font-size: 7pt; text-align: center;"><%=getadCustomSearchIncLngStr("LtxtAdvanced")%></td>
		<td style="width: 8px;"><img src="images/<%=Session("rtl")%>arrows_white.gif" width="8" height="8"></td>
	</tr>
</table>
<table cellpadding="2" cellspacing="0" border="0" id="tblAdSearch" style=" border: 1px solid #FFFFFF; background-color: #0065CE; display: none; position: absolute; z-index: 1; top:100px; left: 100px;" onmouseover="clearTimeout(curAdSearchTimerID);" onmouseout="hideAdSearch();">
	<% do while not rs.eof %>
	<tr onmouseover="this.bgColor='#0075ea';" onmouseout="this.bgColor='';">
		<td style="color: #FFFFFF; font-family: Verdana; font-size: 7pt; cursor: hand; padding-right: 2px; padding-left: 2px; <% If rs.bookmark < rs.recordcount Then %>border-bottom: 1px solid #FFFFFF;<% End If %>" onclick="goAdSearch(<%=rs("ID")%>, <%=ObjID%>);"><nobr><%=rs("Name")%>&nbsp;&nbsp;&nbsp;&nbsp;</nobr></td>
	</tr>
	<% rs.movenext
	loop %>
</table>
<% End If %>
<% addLngPathStr = "searchInc/" %>
<!--#include file="lang/confirmReload.asp" -->
<% If Request.Form("Refresh") <> "" Then
	SelectedValue = CInt(Request("Refresh"))
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKAgentObjectRefreshTime" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@SlpCode") = Session("vendid")
	cmd("@Refresh") = CInt(Request("Refresh"))
	cmd("@ObjType") = "DC"
	cmd("@ObjID") = 0
	cmd.execute()
Else
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKGetAgentObjectRefreshTime" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@SlpCode") = Session("vendid")
	cmd("@ObjType") = "DC"
	cmd("@ObjID") = 0
	set rs = cmd.execute()
	SelectedValue = CInt(rs("Refresh"))
	rs.close
End If %>

<div align="left">
	<table border="0" cellpadding="0" width="93%" id="table3">
		<tr>
			<td align="center">
			<p align="center">
			<input type="button" value="<%=getconfirmReloadLngStr("DtxtUpdate")%>" name="btnReload" style="color: #FFFFFF; font-family: Verdana; font-size: 7pt; border: 1px solid #FFFFFF; background-color: #0065CE; width: 79px" onclick="javascript:doConfRefresh(<%=SelectedValue%>);"></td>
		</tr>
		<tr>
			<td align="center">
			<font size="1" face="Verdana" color="#FFFFFF"><%=getconfirmReloadLngStr("LtxtRefresh")%></font></td>
		</tr>
		<tr>
			<td align="center">
			<select size="1" name="cmbRefresh" style="color: #FFFFFF; font-family: Verdana; font-size: 7pt; border: 1px solid #FFFFFF; background-color: #0065CE" onchange="javascript:doConfRefresh(this.value);">
			<option value="0"><%=getconfirmReloadLngStr("DtxtDisabled")%></option>
			<option <% If SelectedValue = 1 Then %>selected<% End If %> value="1">1 <%=getconfirmReloadLngStr("DtxtMinute")%></option>
			<option <% If SelectedValue = 5 Then %>selected<% End If %> value="5">5 <%=getconfirmReloadLngStr("DtxtMinutes")%></option>
			<option <% If SelectedValue = 10 Then %>selected<% End If %> value="10">10 <%=getconfirmReloadLngStr("DtxtMinutes")%></option>
			<option <% If SelectedValue = 30 Then %>selected<% End If %> value="30">30 <%=getconfirmReloadLngStr("DtxtMinutes")%></option>
			<option <% If SelectedValue = 60 Then %>selected<% End If %> value="60">1 <%=getconfirmReloadLngStr("DtxtHour")%></option>
			</select></td>
		</tr>
	</table>
</div>
<script language="javascript">
var myTimer;
function doConfRefresh(value)
{
	doMyLink('executeConf.asp', 'Type=<%=Request("Type")%>&Refresh=' + value, '_self');
}
<% If SelectedValue <> 0 Then %>
myTimer = setTimeout('doConfRefresh(<%=SelectedValue%>);', <%=SelectedValue*60000%>);
<% End If %>
</script>
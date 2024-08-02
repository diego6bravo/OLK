<% addLngPathStr = "searchInc/" %>
<!--#include file="lang/repReload.asp" -->
<form action="report.asp" method="post" name="frmReload">
<% For each item in Request.Form
If item <> "btnReload" and item <> "Refresh" Then %>
<input type="hidden" name="<%=item%>" value="<%=Request(item)%>">
<% End If
Next
For each item in Request.QueryString
If item <> "Refresh" Then %>
<input type="hidden" name="<%=item%>" value="<%=Request(item)%>">
<% End If
Next %>
<input type="hidden" name="Excell" value="N">
<input type="hidden" name="Refresh" value="">
<div align="left">
	<table border="0" cellpadding="0" width="93%" id="table3">
		<tr>
			<td align="center">
			<p align="center">
			<input type="submit" value="<%=getrepReloadLngStr("DtxtRefresh")%>" name="btnReload" style="color: #FFFFFF; font-family: Verdana; font-size: 7pt; border: 1px solid #FFFFFF; background-color: #0065CE; width: 79"></td>
		</tr>
		<tr>
			<td align="center">
			<font size="1" face="Verdana" color="#FFFFFF"><%=getrepReloadLngStr("LtxtRefresh")%></font></td>
		</tr>
		<tr>
			<td align="center">
			<% If Request.Form("Refresh") <> "" Then
				SelectedValue = CInt(Request("Refresh"))
				set cmd = Server.CreateObject("ADODB.Command")
				cmd.ActiveConnection = connCommon
				cmd.CommandType = &H0004
				cmd.CommandText = "DBOLKAgentObjectRefreshTime" & Session("ID")
				cmd.Parameters.Refresh()
				cmd("@SlpCode") = Session("vendid")
				cmd("@Refresh") = CInt(Request("Refresh"))
				cmd("@ObjType") = "RP"
				cmd("@ObjID") = CInt(Request("rsIndex"))
				cmd.execute()
			Else
				set cmd = Server.CreateObject("ADODB.Command")
				cmd.ActiveConnection = connCommon
				cmd.CommandType = &H0004
				cmd.CommandText = "DBOLKGetAgentObjectRefreshTime" & Session("ID")
				cmd.Parameters.Refresh()
				cmd("@SlpCode") = Session("vendid")
				cmd("@ObjType") = "RP"
				cmd("@ObjID") = CInt(Request("rsIndex"))
				set rs = cmd.execute()
				SelectedValue = CInt(rs("Refresh"))
				rs.close
			End If %>
			<select size="1" name="cmbRefresh" style="color: #FFFFFF; font-family: Verdana; font-size: 7pt; border: 1px solid #FFFFFF; background-color: #0065CE" onchange="javascript:document.frmReload.Refresh.value=this.value;document.frmReload.submit();;">
			<option value="0"><%=getrepReloadLngStr("DtxtDisabled")%></option>
			<option <% If SelectedValue = 1 Then %>selected<% End If %> value="1"><%=getrepReloadLngStr("DtxtMinute")%></option>
			<option <% If SelectedValue = 5 Then %>selected<% End If %> value="5">5 <%=getrepReloadLngStr("DtxtMinutes")%></option>
			<option <% If SelectedValue = 10 Then %>selected<% End If %> value="10">10 <%=getrepReloadLngStr("DtxtMinutes")%></option>
			<option <% If SelectedValue = 30 Then %>selected<% End If %> value="30">30 <%=getrepReloadLngStr("DtxtMinutes")%></option>
			<option <% If SelectedValue = 60 Then %>selected<% End If %> value="60"><%=getrepReloadLngStr("DtxtHour")%></option>
			</select></td>
		</tr>
		<tr>
			<td align="center">
			<font size="1">&nbsp;</font></td>
		</tr>
		<tr>
			<% If HaveVals or rsTop Then %><td align="center">
			<input type="button" value="<%=getrepReloadLngStr("DtxtNew")%>" name="B2" onclick="javascript:<% If userType = "V" Then %>Pic('portal/viewRepVals.asp?rsIndex=<%=Request.Form("rsIndex")%>', 368, 402, 'Yes', 'no')<% ElseIf userType = "C" Then %>document.frmReps.submit();<% End If %>" style="color: #FFFFFF; font-family: Verdana; font-size: 7pt; border: 1px solid #FFFFFF; background-color: #0065CE; width: 79"></td><% End If %>
		</tr>
	</table>
</div>
</form>
<script language="javascript">
function saveRepPdf(Excell)
{
	if (Excell == 'N')
		document.frmReload.action = 'portal/viewRepPDF.asp';
	else
		document.frmReload.action = 'portal/viewReportPDF.asp';
		
	document.frmReload.target = '_blank';
	document.frmReload.Excell.value = Excell;
	document.frmReload.submit();
	document.frmReload.target = '';
	document.frmReload.action = 'report.asp'
}
</script>
<!--#include file="repLegend.asp"-->
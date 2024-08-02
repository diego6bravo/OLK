<% addLngPathStr = "searchInc/" %>
<!--#include file="lang/reload.asp" -->
<form action="<%=strScriptName%>" method="post" name="frmReload">
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
<input type="hidden" name="Refresh" value="<%=Request("Refresh")%>">
<div align="left">
	<table border="0" cellpadding="0" width="93%" id="table3">
		<tr>
			<td align="center">
			<p align="center">
			<input type="submit" value="<%=getreloadLngStr("DtxtUpdate")%>" name="btnReload" style="color: #FFFFFF; font-family: Verdana; font-size: 7pt; border: 1px solid #FFFFFF; background-color: #0065CE; width: 79"></td>
		</tr>
		<tr>
			<td align="center">
			<font size="1" face="Verdana" color="#FFFFFF"><%=getreloadLngStr("LtxtRefresh")%></font></td>
		</tr>
		<tr>
			<td align="center">
			<select size="1" name="cmbRefresh" style="color: #FFFFFF; font-family: Verdana; font-size: 7pt; border: 1px solid #FFFFFF; background-color: #0065CE" onchange="javascript:document.frmReload.Refresh.value=this.value;document.frmReload.submit();;">
			<option value="0"><%=getreloadLngStr("DtxtDisabled")%></option>
			<option <% If Request("Refresh") = "1" Then %>selected<% End If %> value="1"><%=getreloadLngStr("DtxtMinute")%></option>
			<option <% If Request("Refresh") = "5" Then %>selected<% End If %> value="5">5 <%=getreloadLngStr("DtxtMinutes")%></option>
			<option <% If Request("Refresh") = "10" Then %>selected<% End If %> value="10">10 <%=getreloadLngStr("DtxtMinutes")%></option>
			<option <% If Request("Refresh") = "30" Then %>selected<% End If %> value="30">30 <%=getreloadLngStr("DtxtMinutes")%></option>
			<option <% If Request("Refresh") = "60" Then %>selected<% End If %> value="60"><%=getreloadLngStr("DtxtHour")%></option>
			</select></td>
		</tr>
		<tr>
			<td align="center">
			<font size="1">&nbsp;</font></td>
		</tr>
	</table>
</div>
</form>
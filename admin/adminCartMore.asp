<!--#include file="top.asp" -->
<!--#include file="lang/adminCartMore.asp" -->
<% conn.execute("use [" & Session("olkdb") & "]")
sql = "select ShowCSearchTree, ShowCAdSearch from OLKCommon"
set rs = conn.execute(sql) %>
<form method="POST" action="adminsubmit.asp" name="Form1">
	<table border="0" cellpadding="0" width="100%" id="table3">
		<tr>
			<td bgcolor="#E1F3FD"><b><font color="#FFFFFF">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><%=getadminCartMoreLngStr("DtxtCart")%>&nbsp;<%=getadminCartMoreLngStr("LttlCartMore")%></font></b></td>
		</tr>
		<tr>
			<td bgcolor="#F5FBFE">
			<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1">
			</font><font face="Verdana" size="1" color="#4783C5"><%=getadminCartMoreLngStr("LttlCartMoreNote")%></font></p>
			</td>
		</tr>
		<tr>
			<td bgcolor="#F5FBFE">
			<div align="left">
				<table border="0" cellpadding="0" width="100%" id="table6">
					<tr>
						<td width="325" bgcolor="#F7FBFF">
						<img src="images/ganchito.gif"><font color="#4783C5">
						</font><font face="Verdana" size="1" color="#4783C5">
						<input class="noborder" type="checkbox" <% If rs("ShowCAdSearch") = "Y" Then %>checked<% End If %> name="ShowCAdSearch" value="Y" id="ShowCAdSearch"><label for="ShowCAdSearch"><%=getadminCartMoreLngStr("LtxtShowCAdSearch")%></label></font></td>
						<td bgcolor="#F7FBFF">
						&nbsp;</td>
					</tr>
					</table>
			</div>
			</td>
		</tr>
		<tr>
			<td bgcolor="#F5FBFE">
			<div align="left">
				<table border="0" cellpadding="0" width="100%" id="table6">
					<tr>
						<td width="325" bgcolor="#F7FBFF">
						<img src="images/ganchito.gif"><font color="#4783C5">
						</font><font face="Verdana" size="1" color="#4783C5">
						<input class="noborder" type="checkbox" <% If rs("ShowCSearchTree") = "Y" Then %>checked<% End If %> name="ShowCSearchTree" value="Y" id="ShowCSearchTree"><label for="ShowCSearchTree"><%=getadminCartMoreLngStr("LtxtShowCSearchTree")%></label></font></td>
						<td bgcolor="#F7FBFF">
						&nbsp;</td>
					</tr>
					</table>
			</div>
			</td>
		</tr>
		<tr>
			<td>
			<table border="0" cellpadding="0" width="100%" id="table5">
				<tr>
					<td width="77">
					<input type="submit" value="<%=getadminCartMoreLngStr("DtxtSave")%>" name="B1" class="OlkBtn"></td>
					<td><hr color="#0D85C6" size="1"></td>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td>&nbsp;</td>
		</tr>
	</table>
	<input type="hidden" name="submitCmd" value="adminCartMore">
</form>

<!--#include file="bottom.asp" -->
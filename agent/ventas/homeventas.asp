<table border="0" cellpadding="0" bordercolor="#111111" width="100%">
	<tr>
		<td width="100%">
		<table border="0" cellpadding="0" bordercolor="#111111" width="100%">
			<tr>
				<td width="100%">
				<table border="0" cellpadding="0" bordercolor="#111111" width="100%">
					<tr>
						<% If Request("onlyMsg") <> "Y" Then %>
						<td width="50%" valign="top">
						<!--#include file="../taskMonitor.asp"--></td>
						<% End If %>
						<td valign="top" rowspan="2">
						<!--#include file="../messages/messages.asp" --></td>
					</tr>
					<% If Request("onlyMsg") <> "Y" Then %>
					<tr>
						<td width="50%" valign="top">
						<!--#include file="../news/top2news.asp" --><% doTopNews %>
						</td>
					</tr>
					<% End If %>
				</table>
				</td>
			</tr>
		</table>
		</td>
	</tr>
</table>


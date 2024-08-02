<%
sql = "select  case when exists(select 'A' from OLKCUFD where TableID = 'OCRD' and Active = 'Y') Then 'Y' Else 'N' End EnableSDK "
set rd = conn.execute(sql)
EnableSDK = rd("EnableSDK") = "Y"
 %>
<table style="width: 100%;" cellspacing="0" cellpadding="0">
				<tr>
								<td style="width: 20%; text-align: center;">
								<input type="image" name="btnMain" src="activity/act_main.jpg" width="40" height="40" border="0" <% If Request("cmd") = "newClientUDF" Then %>onclick="return valFrm();"<% End If %>></td>
								<td style="width: 20%; text-align: center;">
								<input type="image" name="btnGeneral" src="client/cl_main_more.jpg" width="40" height="40" border="0" <% If Request("cmd") = "newClientUDF" Then %>onclick="return valFrm();"<% End If %>></td>
								<td style="width: 20%; text-align: center;">
								<input type="image" name="btnAddress" src="activity/act_address.jpg" width="40" height="40" border="0" <% If Request("cmd") = "newClientUDF" Then %>onclick="return valFrm();"<% End If %>></td>
								<td style="width: 20%; text-align: center;">
								<input type="image" name="btnContacts" src="client/cl_contacts.jpg" width="40" height="40" border="0" <% If Request("cmd") = "newClientUDF" Then %>onclick="return valFrm();"<% End If %>></td>
								<td style="width: 20%; text-align: center;"><% If EnableSDK Then %>
								<input type="image" name="btnUDF" src="activity/act_udf.jpg" width="40" height="40" border="0" <% If Request("cmd") = "newClientUDF" Then %>onclick="return valFrm();"<% End If %>><% Else %>&nbsp;<% End If %></td>
				</tr>
</table>

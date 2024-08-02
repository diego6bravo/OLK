<script language="javascript" src="activity/addActivity.js"></script>
<%
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetActivityMenuData" & Session("ID")
cmd.Parameters.Refresh()
cmd("@ID") = Session("ActRetVal")
If Session("ActReadOnly") Then cmd("@ReadOnly") = "Y"
set rd = cmd.execute()

Action = rd("Action")
EnableSDK = rd("EnableSDK") = "Y"

If Request("cmd") = "activity" and Request.Form.Count > 0 Then
	Action = Request("Action")
End If


set rsdf = Server.CreateObject("ADODB.RecordSet")
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetUDFSystemCols" & Session("ID")
cmd.Parameters.Refresh
cmd("@LanID") = Session("LanID")
cmd("@UserType") = userType
cmd("@TableID") = "OCLG"
cmd("@OP") = "P"
rsdf.open cmd, , 3, 1


rsdf.Filter = "FieldID = -32"
showCardCode = not rsdf.eof
rsdf.Filter = "FieldID = -33"
showCardName = not rsdf.eof


 %>
<table style="width: 100%;" cellspacing="0" cellpadding="0">
				<tr>
								<td style="width: 20%; text-align: center;">
								<input type="image" name="btnMain" id="btnMain" src="activity/act_main.jpg" width="40" height="40" border="0"></td>
								<td style="width: 20%; text-align: center;">
								<input type="image" name="btnGeneral" id="btnGeneral" src="activity/act_general.jpg" width="40" height="40" border="0"></td>
								<% rsdf.Filter = "FieldID = -27 or FieldID = -28 or FieldID = -29 or FieldID = -30 or FieldID = -31 "
								If Not rsdf.Eof Then %>
								<td style="width: 20%; text-align: center;"><% If Action = "M" Then %>
								<input type="image" name="btnAddress" id="btnAddress" src="activity/act_address.jpg" width="40" height="40" border="0"><% Else %>&nbsp;<% End If %></td><% End If %>
								<% rsdf.Filter = "FieldID = -26"
							    If Not rsdf.Eof Then %>
								<td style="width: 20%; text-align: center;">
								<input type="image" name="btnContent" id="btnContent" src="activity/act_content.jpg" width="40" height="40" border="0"></td><% End If %>
								<td style="width: 20%; text-align: center;"><% If EnableSDK Then %>
								<input type="image" name="btnUDF" src="activity/act_udf.jpg" width="40" height="40" border="0"><% Else %>&nbsp;<% End If %></td>
				</tr><% If showCardCode or showCardName Then %>
				<tr>
								<td colspan="5">
								<table cellpadding="0" cellspacing="0" border="0">
												<tr>
																<td align="right">
																<a href='operaciones.asp?cmd=datos&amp;card=<%=CleanItem(rd("CardCode"))%>'>
																<img border="0" src='images/<%=Session("rtl")%>flechaselec.gif'></a></td>
																<td style="font-family: Verdana; font-size: xx-small; "><strong><% If showCardCode Then %><%=rd("CardCode")%><% End If %><% If showCardCode and showCardName Then %> - <% End If %><% If showCardName Then %><%=rd("CardName")%><% End If %></strong></td>
												</tr>
								</table>
								</td>
				</tr>
				<% End If %>
</table>

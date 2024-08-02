<% addLngPathStr = "activity/" %>
<!--#include file="lang/activityAddress.asp" -->

<head>
<style type="text/css">
.style1 {
				font-family: Verdana;
				font-size: xx-small;
}
.style5 {
				font-family: Verdana;
				font-size: xx-small;
				background-color: #75ACFF;
}
</style>
</head>

<% 
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetActivityAddressData" & Session("ID")
cmd.Parameters.Refresh()
cmd("@ID") = Session("ActRetVal")
If Session("ActReadOnly") Then cmd("@ReadOnly") = "Y"
set rs = cmd.execute()
ReadOnly = Session("ActReadOnly")

ClgCode = rs("ClgCode")

If Request.Form.Count = 0 Then
	Country = rs("Country")
	State = rs("State")
	City = rs("City")
	Street = rs("Street")
	Room = rs("Room")
Else
	Country = Request("Country")
	State = Request("State")
	City = Request("City")
	Street = Request("Street")
	Room = Request("Room")
End If

 %>
<div align="center">
	<center>
	<table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111" bgcolor="#9BC4FF">
	<form name="frmAddress" method="post" action="activity/actSubmit.asp">
		<input type="hidden" name="cmd" value="address">
		<tr>
			<td width="100%" align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
				<table cellpadding="0" border="0">
					<tr>
						<td><img src="images/icon_activity_<% If Not IsNull(ClgCode) Then %>S<% Else %>O<% End If %>.gif"></td>
						<td><b><font face="Verdana" size="1"><%=getactivityAddressLngStr("DtxtActivity")%>&nbsp;#<% If Not IsNull(ClgCode) Then Response.Write ClgCode Else Response.Write Session("ActRetVal") %>&nbsp;-&nbsp;<%=getactivityAddressLngStr("DtxtAddress")%></font></b></td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td width="100%">
			<!--#include file="activityMenu.asp"--></td>
		</tr>
		<tr>
			<td>
			<table style="width: 100%">
				<% rsdf.Filter = "FieldID = -27"
			        If Not rsdf.Eof Then %>
				<tr>
					<td class="style5"><strong><% If rsdf("AlterName") <> "" Then %><%=rsdf("AlterName")%><% Else %><%=getactivityAddressLngStr("DtxtCountry")%><% End If %></strong></td>
					<td>
					<% 
					set cmd = Server.CreateObject("ADODB.Command")
					cmd.ActiveConnection = connCommon
					cmd.CommandType = &H0004
					cmd.CommandText = "DBOLKGetCountries" & Session("ID")
					cmd.Parameters.Refresh()
					cmd("@LanID") = Session("LanID")
					set rd = cmd.execute() %>
					<select <% If ReadOnly Then %>disabled<% End If %> class="input" size="1" name="Country" style="font-size: 10px; font-family: Verdana" onchange="javascript:changeCountry();">
					<option></option>
					<% do while not rd.eof %>
					<option <% If rd("Code") = Country Then %>selected<% End If %>="" value='<%=rd("Code")%>'><%=myHTMLEncode(rd("name"))%></option>
					<% rd.movenext
					loop
					%></select></td>
				</tr>
				<% End If %>
				<% rsdf.Filter = "FieldID = -28"
			        If Not rsdf.Eof Then %>
				<tr>
					<td class="style5"><strong><% If rsdf("AlterName") <> "" Then %><%=rsdf("AlterName")%><% Else %><%=getactivityAddressLngStr("DtxtState")%><% End If %></strong></td>
					<td><select <% If ReadOnly Then %>disabled<% End If %> class="input" size="1" name="State" style="font-size: 10px; font-family: Verdana">
					<option></option>
					<% 
					If Country <> "" Then
					
					set cmd = Server.CreateObject("ADODB.Command")
					cmd.ActiveConnection = connCommon
					cmd.CommandType = &H0004
					set rd = Server.CreateObject("ADODB.RecordSet")
					cmd.CommandText = "DBOLKGetCountryStates" & Session("ID")
					cmd.Parameters.Refresh()
					cmd("@LanID") = Session("LanID")
					cmd("@Code") = Country
					rd.open cmd, , 3, 1
					do while not rd.eof %>
					<option <% If rd("Code") = State Then %>selected<% End If %>="" value='<%=rd("Code")%>'><%=myHTMLEncode(rd("name"))%></option>
					<% rd.movenext
					loop
					End If
					%></select></td>
				</tr>
				<% End If %>
				<% rsdf.Filter = "FieldID = -29"
			        If Not rsdf.Eof Then %>
				<tr>
					<td class="style5"><strong><% If rsdf("AlterName") <> "" Then %><%=rsdf("AlterName")%><% Else %><%=getactivityAddressLngStr("DtxtCity")%><% End If %></strong></td>
					<td><input <% If ReadOnly Then %>readonly<% End If %> class="input" type="text" name="City" size="25" value='<%=myHTMLEncode(city)%>' onkeydown="return chkMax(event, this, 100);" style="width: 222px"></td>
				</tr>
				<% End If %>
				<% rsdf.Filter = "FieldID = -30"
			        If Not rsdf.Eof Then %>
				<tr>
					<td valign="top" class="style5"><strong><% If rsdf("AlterName") <> "" Then %><%=rsdf("AlterName")%><% Else %><%=getactivityAddressLngStr("DtxtStreet")%><% End If %></strong></td>
					<td><textarea <% If ReadOnly Then %>readonly<% End If %> rows="3" name="Street" cols="20" class="input" style="width: 96%"><%=myHTMLEncode(Street)%></textarea></td>
				</tr>
				<% End If %>
				<% rsdf.Filter = "FieldID = -31"
			        If Not rsdf.Eof Then %>
				<tr>
					<td class="style5"><strong><% If rsdf("AlterName") <> "" Then %><%=rsdf("AlterName")%><% Else %><%=getactivityAddressLngStr("DtxtRoom")%><% End If %></strong></td>
					<td><input <% If ReadOnly Then %>readonly<% End If %> class="input" type="text" name="Room" size="35" value='<%=myHTMLEncode(Room)%>' onkeydown="return chkMax(event, this, 50);" style="width: 222px"></td>
				</tr>
				<% End If %>
				</table>
			</td>
		</tr>
		<!--#include file="activityBottom.asp"-->
	</form>
	</table>
	</center>
</div>
<script type="text/javascript">
function changeCountry()
{
document.frmAddress.action = 'operaciones.asp';
document.frmAddress.cmd.value = 'activityAddress';
document.frmAddress.State.selectedIndex = 0;
document.frmAddress.submit();
}
</script>

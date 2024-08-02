<% addLngPathStr = "activity/" %>
<!--#include file="lang/activityMain.asp" -->
<head>
<style type="text/css">
.style1 {
				font-family: Verdana;
				font-size: xx-small;
}
.style2 {
				font-family: Verdana;
}
.style3 {
				font-size: xx-small;
}
.style4 {
				background-color: #75ACFF;
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
cmd.CommandText = "DBOLKGetActivityMainData" & Session("ID")
cmd.Parameters.Refresh()
cmd("@ID") = Session("ActRetVal")
If Session("ActReadOnly") Then cmd("@ReadOnly") = "Y"
set rs = cmd.execute()

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004

ReadOnly = Session("ActReadOnly")

ClgCode = rs("ClgCode")
If Request.Form.Count = 0 Then
	CntctCode = rs("CntctCode")
	Tel = rs("Tel")
	Action = rs("Action")
	CntctType = rs("CntctType")
	CntctSbjct = rs("CntctSbjct")
	AttendUser = rs("AttendUser")
	Priority = rs("Priority")
	Details = rs("Details")
	personal = rs("personal") = "Y"
	Closed = rs("Closed") = "Y"
Else
	CntctCode = Request("CntctCode")
	Tel = Request("Tel")
	Action = Request("Action")
	CntctType = Request("CntctType")
	CntctSbjct = Request("CntctSbjct")
	AttendUser = Request("AttendUser")
	Priority = Request("Priority")
	Details = Request("Details")
	personal = Request("personal") = "Y"
	Closed = Request("Closed") = "Y"
End If

 %>
<div align="center">
	<center>
	<table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111" bgcolor="#9BC4FF">
        <tr>
          <td width="100%" align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
          <table cellpadding="0" border="0">
			<tr>
				<td><img src="images/icon_activity_<% If Not IsNull(ClgCode) Then %>S<% Else %>O<% End If %>.gif"></td>
				<td><b><font face="Verdana" size="1"><%=getactivityMainLngStr("DtxtActivity")%>&nbsp;#<% If Not IsNull(ClgCode) Then Response.Write ClgCode Else Response.Write Session("ActRetVal") %></font></b></td>
			</tr>
			</table>
          </td>
        </tr>
		<form name="frmActivity" method="post" action="activity/actSubmit.asp">
		<input type="hidden" name="cmd" value="main">
        <tr>
          <td width="100%">
          <!--#include file="activityMenu.asp"--></td>
        </tr>
		<tr>
			<td>
			<table style="width: 100%">
				<% rsdf.Filter = "FieldID = -1"
			        If Not rsdf.Eof Then %>
				<tr>
				<% 
				
		        set cmd = Server.CreateObject("ADODB.Command")
				cmd.ActiveConnection = connCommon
				cmd.CommandType = &H0004
				cmd.CommandText = "DBOLKGetBPContacts" & Session("ID")
				cmd.Parameters.Refresh()
				cmd("@LanID") = Session("LanID")
				cmd("@CardCode") = Session("UserName")
		        set rd = cmd.execute() %>
				<td class="style5"><strong><% If rsdf("AlterName") <> "" Then %><%=rsdf("AlterName")%><% Else %><%=getactivityMainLngStr("DtxtContact")%><% End If %></strong></td>
				<td><select <% If ReadOnly Then %>disabled<% End If %> class="input" size="1" name="CntctCode" style="font-size: 10px; font-family: Verdana" onchange="javascript:doChangeContact();"><% do while not rd.eof %>
				<option <% If CStr(rd("CntctCode")) = CStr(CntctCode) Then %>selected<% End If %> value='<%=rd("cntctcode")%>'><%=myHTMLEncode(rd("name"))%></option>
				<% rd.movenext
				loop
				%></select></td>
				</tr>
				<% End If %>
				<% rsdf.Filter = "FieldID = -2"
			        If Not rsdf.Eof Then %>
				<tr>
					<td class="style5"><strong><% If rsdf("AlterName") <> "" Then %><%=rsdf("AlterName")%><% Else %><%=getactivityMainLngStr("DtxtPhone")%><% End If %></strong></td>
					<td><input <% If ReadOnly Then %>readonly<% End If %> class="input" type="text" name="Tel" size="20" value="<%=myHTMLEncode(Tel)%>" maxlength="20"></td>
				</tr>
				<% End If %>
				<% rsdf.Filter = "FieldID = -3"
			        If Not rsdf.Eof Then %>
				<tr>
					<td class="style5"><strong><% If rsdf("AlterName") <> "" Then %><%=rsdf("AlterName")%><% Else %><%=getactivityMainLngStr("DtxtActivity")%><% End If %></strong></td>
					<td><select <% If ReadOnly Then %>disabled<% End If %> size="1" class="input" name="Action" onchange="javascript:changeAction();">
					<option <% If Action = "C" Then %>selected<% End If %> value="C"><%=getactivityMainLngStr("DtxtConv")%></option>
					<option <% If Action = "M" Then %>selected<% End If %> value="M"><%=getactivityMainLngStr("DtxtMeeting")%></option>
					<option <% If Action = "E" Then %>selected<% End If %> value="E"><%=getactivityMainLngStr("DtxtNote")%></option>
					<option <% If Action = "O" Then %>selected<% End If %> value="O"><%=getactivityMainLngStr("DtxtOther")%></option>
					<option <% If Action = "T" Then %>selected<% End If %> value="T"><%=getactivityMainLngStr("DtxtTask")%></option>
					</select></td>
				</tr>
				<% End If %>
				<% rsdf.Filter = "FieldID = -4"
			        If Not rsdf.Eof Then %>
				<tr>
					<td class="style5"><strong><% If rsdf("AlterName") <> "" Then %><%=rsdf("AlterName")%><% Else %><%=getactivityMainLngStr("DtxtType")%><% End If %></strong></td>
					<td>
					<% 
					set rd = Server.CreateObject("ADODB.RecordSet")
					cmd.CommandText = "DBOLKGetActTypes" & Session("ID")
					cmd.Parameters.Refresh()
					cmd("@LanID") = Session("LanID")
					rd.open cmd, , 3, 1																				
					 %>
					<select <% If ReadOnly Then %>disabled<% End If %> class="input" size="1" name="CntctType" style="font-size: 10px; font-family: Verdana" onchange="javascript:changeType();">
					<% do while not rd.eof %>
					<option <% If CStr(rd("Code")) = CStr(CntctType) Then %>selected<% End If %> value='<%=rd("Code")%>'><%=myHTMLEncode(rd("name"))%></option>
					<% rd.movenext
					loop
					%></select></td>
				</tr>
				<% End If %>
				<% rsdf.Filter = "FieldID = -5"
			        If Not rsdf.Eof Then %>
				<tr>
					<td class="style5"><strong><% If rsdf("AlterName") <> "" Then %><%=rsdf("AlterName")%><% Else %><%=getactivityMainLngStr("DtxtSubject")%><% End If %></strong></td>
					<td>
					<% 
					
					set cmd = Server.CreateObject("ADODB.Command")
					cmd.ActiveConnection = connCommon
					cmd.CommandType = &H0004
					set rd = Server.CreateObject("ADODB.RecordSet")
					cmd.CommandText = "DBOLKGetActivitySubjects" & Session("ID")
					cmd.Parameters.Refresh()
					cmd("@LanID") = Session("LanID")
					cmd("@Type") = CntctType
					rd.open cmd, , 3, 1 %>
					<select <% If ReadOnly Then %>disabled<% End If %> class="input" size="1" name="CntctSbjct" style="font-size: 10px; font-family: Verdana">
					<option></option>
					<% do while not rd.eof %>
					<option <% If CStr(rd("Code")) = CStr(CntctSbjct) Then %>selected<% End If %> value='<%=rd("Code")%>'><%=myHTMLEncode(rd("name"))%></option>
					<% rd.movenext
					loop
					%></select></td>
				</tr>
				<% End If %>
				<% rsdf.Filter = "FieldID = -6"
			        If Not rsdf.Eof Then %>
				<tr>
					<td class="style5"><strong><% If rsdf("AlterName") <> "" Then %><%=rsdf("AlterName")%><% Else %><%=getactivityMainLngStr("LtxtAsignedTo")%><% End If %></strong></td>
					<td>
				<% 	set cmd = Server.CreateObject("ADODB.Command")
					cmd.ActiveConnection = connCommon
					cmd.CommandType = &H0004
					cmd.CommandText = "DBOLKGetSysUsers" & Session("ID")
					cmd.Parameters.Refresh()
					cmd("@LanID") = Session("LanID")
					set rd = cmd.execute() %>
					<select <% If ReadOnly Then %>disabled<% End If %> class="input" size="1" name="AttendUser" style="font-size: 10px; font-family: Verdana"><% do while not rd.eof %>
					<option <% If CStr(rd("Code")) = CStr(AttendUser) Then %>selected<% End If %> value='<%=rd("Code")%>'><%=myHTMLEncode(rd("Name"))%></option>
					<% rd.movenext
					loop
					%></select></td>
				</tr>
				<% End If %>
				<% rsdf.Filter = "FieldID = -7"
			        If Not rsdf.Eof Then %>
				<tr>
					<td class="style5"><strong><% If rsdf("AlterName") <> "" Then %><%=rsdf("AlterName")%><% Else %><%=getactivityMainLngStr("DtxtPriority")%><% End If %></strong></td>
					<td><select <% If ReadOnly Then %>disabled<% End If %> class="input" size="1" name="Priority" style="font-size: 10px; font-family: Verdana">
					<option <% If Priority = "L" Then %>selected<% End If %> value="L"><%=getactivityMainLngStr("DtxtLow")%></option>
					<option <% If Priority = "N" Then %>selected<% End If %> value="N"><%=getactivityMainLngStr("DtxtNormal")%></option>
					<option <% If Priority = "H" Then %>selected<% End If %> value="H"><%=getactivityMainLngStr("DtxtHigh")%></option>
					</select></td>
				</tr>
				<% End If %>
				<% rsdf.Filter = "FieldID = -8"
			        If Not rsdf.Eof Then %>
				<tr>
					<td colspan="2" class="style5"><strong><% If rsdf("AlterName") <> "" Then %><%=rsdf("AlterName")%><% Else %><%=getactivityMainLngStr("DtxtCommentaries")%><% End If %></strong></td>
				</tr>
				<tr>
					<td colspan="2"><textarea <% If ReadOnly Then %>readonly<% End If %> rows="6" name="Details" cols="20" class="input" style="width: 96%" MaxLength="60"><%=myHTMLEncode(Details)%></textarea></td>
				</tr>
				<% End If %>
				<% rsdf.Filter = "FieldID = -9"
			        If Not rsdf.Eof Then %>
				<tr>
					<td class="style5" colspan="2"><font size="1" face="Verdana"><strong><span class="style2"><span class="style3"><input  <% If ReadOnly Then %>disabled<% End If %> type="checkbox" style="background: background-image; border: 0px solid" name="personal" id="personal" value="Y" <% If personal Then %> checked<% End If %>></span></span></strong></font><span class="style1"><label for="personal"><strong><%=getactivityMainLngStr("DtxtPersonal")%></strong></label></span></td>
				</tr>
				<% End If %>
				<% rsdf.Filter = "FieldID = -10"
			        If Not rsdf.Eof Then %>
				<tr>
					<td class="style5" colspan="2"><span class="style2"><span class="style3"><strong><input <% If ReadOnly Then %>disabled<% End If %> type="checkbox" name="Closed" id="Closed" value="Y" <% If Closed Then %>checked<% End If %> style="background: background-image; border: 0px solid"></strong></span></span><span class="style1"><strong><label for="Closed"><%=getactivityMainLngStr("DtxtClosed")%></label></strong></span></td>
				</tr>
				<% End If %>
			</table>
			</td>
		</tr>
		<input type="hidden" name="changeContact" value="">
		<!--#include file="activityBottom.asp"-->
						</form>
		</table>
		</center></div>
<script type="text/javascript">
<!--
function valFrm()
{
	return true;
}
function doChangeContact()
{
	document.frmActivity.changeContact.value = 'Y';
	document.getElementById('btnMain').click();
}
//-->
</script>
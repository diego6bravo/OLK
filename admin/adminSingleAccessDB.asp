<!--#include file="top.asp" -->
<!--#include file="lang/adminSingleAccessDB.asp" --><head>
<style type="text/css">
.style1 {
	text-align: center;
	font-weight: bold;
	background-color: #E1F3FD;
}
.style2 {
	background-color: #F5FBFE;
}
.style3 {
	background-color: #E1F3FD;
}
.style4 {
	font-weight: bold;
	background-color: #E1F3FD;
}
.style5 {
	text-align: center;
	background-color: #F5FBFE;
}
.style6 {
	text-align: center;
	background-color: #E1F3FD;
}
.style7 {
	text-align: center;
	background-color: #F3FBFE;
}
.style8 {
	background-color: #F3FBFE;
}
</style>

</head>
<%
UserName = Request("UserName")

If Request.Form.Count > 0 Then
	If Request("dbID") <> "" Then
		arrID = Split(Request("dbID"), ", ")
		For i = 0 to UBound(arrID)
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "OLKSetDBAgentData"
			cmd.Parameters.Refresh
			dbID = arrID(i)
			cmd("@UserName") = UserName
			cmd("@dbID") = dbID
			
			slpCode = Request("SlpCode" & dbID)
			If slpCode <> "" Then
				cmd("@SlpCode") = slpCode
				cmd("@Access") = Request("Access" & dbID)
				cmd("@WhsCode") = Request("WhsCode" & dbID)
			End If
			cmd.execute()
		Next
	End If
	If Request("btnSave") <> "" Then Response.Redirect "adminSingleAccess.asp"
End If

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "OLKAgentDBList"
cmd.Parameters.Refresh
cmd("@UserName") = UserName
set rs = Server.CreateObject("ADODB.RecordSet")
rs.open cmd, , 3, 1

set ra = Server.CreateObject("ADODB.RecordSet")
set cmda = Server.CreateObject("ADODB.Command")
cmda.ActiveConnection = connCommon
cmda.CommandType = &H0004
cmda.CommandText = "OLKGetDBAgents"
cmda.Parameters.Refresh

%>
<script type="text/javascript">
<!--
var UserName = '<%=Replace(UserName, "'", "\'")%>';

function copyAut(id, Access)
{
	var SlpCode = document.getElementById('SlpCode' + id).value;
	var winleft = (screen.width - 400) / 2;
	var winUp = (screen.height - 300) / 2;
	OpenWin = this.open('adminAgentsAutCopy.asp?UserName=' + escape(UserName) + '&dbID=' + id + '&SlpCode=' + SlpCode+ '&pop=Y&Access=' + Access, 'CopyAut', "toolbar=no,menubar=no,location=no,left="+winleft+",top="+winUp+",scrollbars=1,resizable=0, width=400,height=300");
}

function changeAccess(value, id)
{
	//document.getElementById('btnEditGroups' + SlpCode).style.display = value == 'D' ? 'none' : '';
	document.getElementById('lnkAut' + id).style.display = value == 'D' ? 'none' : '';
	if (document.getElementById('Aut' + id).style.display != 'none') goUserAut(id);
}

function expandUserAut(id)
{
	if (document.getElementById('Aut' + id).style.display == 'none' && document.getElementById('Access' + id).value != 'D')
	{
		showUserAut(id, true);
		if (document.getElementById('iFrame' + id).src == '')
		{
			goUserAut(id);
		}
	}
	else
	{
		showUserAut(id, false);
	}
}

function goUserAut(id)
{
	var SlpCode = document.getElementById('SlpCode' + id).value;
	document.getElementById('iFrame' + id).src = 'adminAgentsAuthorization.asp?UserName=' + escape(UserName) + '&dbID=' + id + '&SlpCode=' + SlpCode + '&parent=Y&Access=' + document.getElementById('Access' + id).value;
}

function showUserAut(id, Show)
{
	document.getElementById('Aut' + id).style.display = Show ? '' : 'none';
	document.getElementById('lnkAut' + id).innerHTML = Show ? '[-]' : '[+]';	
}

function valFrm()
{
	if (document.frmDB.dbID)
	{
		if (document.frmDB.dbID.length)
		{
			for (var i = 0;i<document.frmDB.dbID.length;i++)
			{
				if (!checkData(document.frmDB.dbID[i].value)) return false;
			}
		}
		else
		{
			return checkData(document.frmDB.dbID.value);
		}
	}
	return true;
}
function checkData(dbID)
{
	if (document.getElementById('Access' + dbID).value != 'D' && document.getElementById('SlpCode' + dbID).selectedIndex == 0)
	{
		alert('<%=getadminSingleAccessDBLngStr("LtxtValSlpSel")%>');
		document.getElementById('SlpCode' + dbID).focus();
		return false;
	}
	return true;
}
//-->
</script>
<table border="0" cellpadding="0" width="100%">
	<tr>
		<td height="15"></td>
	</tr>
	<form method="POST" action="adminSingleAccessDB.asp" name="frmDB" onsubmit="return valFrm();">
	<input type="hidden" name="UserName" value="<%=Server.HTMLEncode(UserName)%>">
	<tr>
		<td bgcolor="#E1F3FD"><b><font color="#FFFFFF">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><%=getadminSingleAccessDBLngStr("LtxtTitle")%> </font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1" color="#3066E4"> </font>
		<font face="Verdana" size="1" color="#4783C5"><%=getadminSingleAccessDBLngStr("LtxtNote")%> </font></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><font face="Verdana" size="1" color="#3066E4"> </font>
		<font face="Verdana" size="1" color="#4783C5"><b><%=getadminSingleAccessDBLngStr("DtxtUser")%>:</b> <%=UserName%> </font></td>
	</tr>
	<tr>
		<td>
		<table cellpadding="0" border="0" width="100%"> 
			<tr> 
				<td class="style1" style="width: 20px">&nbsp;</td> 
				<td class="style1"> 
				<font face="Verdana" size="1" color="#31659C"><%=getadminSingleAccessDBLngStr("DtxtDB")%></font></td> 
				<td class="style1"> 
				<font face="Verdana" size="1" color="#31659C"><%=getadminSingleAccessDBLngStr("DtxtCompany")%></font></td> 
				<td class="style1"> 
				<font face="Verdana" size="1" color="#31659C"><%=getadminSingleAccessDBLngStr("DtxtAgent")%></font></td> 
				<td width="110" class="style1"> 
				<font face="Verdana" size="1" color="#31659C"><%=getadminSingleAccessDBLngStr("DtxtAccess")%></font></td> 
				<td class="style1" style="width: 200px"> 
				<font face="Verdana" size="1" color="#31659C"><%=getadminSingleAccessDBLngStr("DtxtWarehouse")%></font></td> 
			</tr> 
			<% If Not rs.Eof Then
			do while not rs.eof
			
			dbID = rs("ID")
			cmda("@dbID") = dbID 
			cmda("@dbName") = rs("dbName")
			If Not IsNull(rs("SlpCode")) Then cmda("@SlpCode") = rs("SlpCode") Else cmda("@SlpCode") = -9
			set ra = cmda.execute()
			
			whsCode = ""
			access = "D"
			If Not IsNull(rs("SlpCode")) Then
				set cmd = Server.CreateObject("ADODB.Command")
				cmd.ActiveConnection = connCommon
				cmd.CommandType = &H0004
				cmd.CommandText = "OLKGetDBAgentData"
				cmd.Parameters.Refresh
				cmd("@dbID") = dbID 
				cmd("@dbName") = rs("dbName")
				cmd("@SlpCode") = rs("SlpCode")
				set rd = Server.CreateObject("ADODB.RecordSet")
				set rd = cmd.execute()
				If Not rd.Eof Then
					whsCode = rd("WhsCode")
					access = rd("Access")
				Else
					whsCode = "##"
					access = "D"
				End If
			End If
			

			set rw = Server.CreateObject("ADODB.RecordSet")
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKGetWarehouses" & dbID 
			cmd.Parameters.Refresh()
			cmd("@LanID") = Session("LanID")
			rw.open cmd, , 3, 1
			%>
			<tr bgcolor="#F3FBFE"> 
				<td class="style2" style="width: 20px"><input type="hidden" name="dbID" value="<%=dbID%>">
				<font face="Verdana" size="1" color="#3366CC">
				<a href="#" id="lnkAut<%=dbID%>" onclick="javascript:expandUserAut(<%=dbID%>);" style="text-decoration: none;<% If access = "D" Then %>display: none;<% End If %> ">[+]</a></font></td> 
				<td class="style2"> 
				<font face="Verdana" size="1" color="#4783C5"><%=rs("dbName")%></font></td> 
				<td class="style2"> 
				<font face="Verdana" size="1" color="#4783C5"><%=rs("cmpName")%></font></td> 
				<td class="style2"> 
				<select name="SlpCode<%=dbID%>" id="SlpCode<%=dbID%>" size="1" class="input" style="width: 100%;"><option></option><%
				do while not ra.eof
				%><option <% If ra("SlpCode") = rs("SlpCode") Then %>selected<% End If %> value="<%=ra("slpcode")%>"><%=myHTMLEncode(ra("slpname"))%></option><% ra.movenext
				loop %>
				</select></td>
				<td width="110" class="style2"> 
				<select name="Access<%=dbID%>" id="Access<%=dbID%>" size="1" class="input" onchange="changeAccess(this.value,<%=dbID%>);"> 
				<option value="D"><%=getadminSingleAccessDBLngStr("DtxtDisabled")%></option> 
				<option <% If access = "U" Then %>selected<% End If %> value="U"><%=getadminSingleAccessDBLngStr("DtxtUser")%></option> 
				<option <% If access = "P" Then %>selected<% End If %> value="P"><%=getadminSingleAccessDBLngStr("LtxtSuperUser")%></option> 
				</select></td> 
				<td class="style2" style="width: 200px"> 
				<select name="WhsCode<%=dbID%>" size="1" class="input" style="width: 100%;">
				<option value="##"><%=getadminSingleAccessDBLngStr("DtxtDefault")%></option>
				<% do while not rw.eof %>
				<option <% If whsCode = rw("WhsCode") Then %>selected<% End If %> value="<%=rw("WhsCode")%>"><%=myHTMLEncode(rw("WhsName"))%></option>
				<% rw.movenext
				loop %>
				</select></td> 
			</tr> 
			<tr bgcolor="#F3FBFE" id="Aut<%=dbID%>" style="display: none"> 
				<td class="style2" colspan="10"> 
				<iframe id="iFrame<%=dbID%>" width="100%" height="350" scrolling="yes"></iframe></td> 
			</tr> 
			<% rs.movenext
			loop
			else %>
			<tr bgcolor="#F3FBFE"> 
				<td class="style2" colspan="6" align="center"><font face="Verdana" size="1" color="#4783C5"><%=getadminSingleAccessDBLngStr("LtxtNoDB")%></font></td> 
			</tr> 
			<% End If %>
		</table> 
		</td>
	</tr>
	<tr>
		<td>
			<table cellpadding="0" border="0" width="100%">
				<td width="75"><input type="submit" value="<%=getadminSingleAccessDBLngStr("DtxtApply")%>" name="btnApply" class="OlkBtn"></td>
				<td width="75"><input type="submit" value="<%=getadminSingleAccessDBLngStr("DtxtSave")%>" name="btnSave" class="OlkBtn"></td>
				<td><hr color="#0D85C6" size="1"></td>
				<td width="75"><input type="button" value="<%=getadminSingleAccessDBLngStr("DtxtCancel")%>" name="btnCancel" class="OlkBtn" onclick="window.location.href='adminSingleAccess.asp';"></td>
			</table>
		</td>
	</tr>
	</form>
</table>
<!--#include file="bottom.asp" -->
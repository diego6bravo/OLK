<!--#include file="top.asp" -->
<!--#include file="lang/adminBranchs.asp" -->

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style type="text/css">
.style1 {
	text-align: center;
}
</style>
</head>

<% 
conn.execute("use [" & Session("OLKDB") & "]")

GetQuery rs, 1, null, null

set rd = server.createobject("ADODB.RecordSet")
GetAdminQuery rd, 3, "L", null
%>
<script language="javascript">
function valAddBranch()
{
	if (document.frmAddBranch.BranchName.value == '')
	{
		alert("<%=getadminBranchsLngStr("LtxtValBranchNam")%>");
		document.frmAddBranch.BranchName.focus();
		return false;
	}
	else if (document.frmAddBranch.WhsCode.selectedIndex == 0)
	{
		alert("<%=getadminBranchsLngStr("LtxtValWhs")%>");
		document.frmAddBranch.WhsCode.focus();
		return false;
	}
	return true;
}
</script>
<br>
<table border="0" cellpadding="0" width="100%" id="table6">
	<form method="POST" action="submitBranchs.asp" name="frmAddBranch" onsubmit="return valAddBranch();">
	<input type="hidden" name="BranchNameTrad">
	<tr>
		<td bgcolor="#E1F3FD"><b><font size="1" color="#31659C" face="Verdana">
		<%=getadminBranchsLngStr("LttlAddBranch")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif">&nbsp;<font color="#4783C5" face="Verdana" size="1"><%=getadminBranchsLngStr("LttlAddBranchNote")%></font></td>
	</tr>
	<tr>
		<td>
		<table cellpadding="0">
			<tr>
				<td style="width: 100px" bgcolor="#E1F3FD"><font color="#31659C" face="Verdana" size="1">
				<strong><%=getadminBranchsLngStr("DtxtName")%></strong></font></td>
				<td style="width: 400px" bgcolor="#F5FBFE"><input type="text" name="BranchName" size="20" style="width: 100%; " class="input" onkeydown="return chkMax(event, this, 50);"></td>
				<td bgcolor="#F5FBFE"><a href="javascript:doFldTrad('Branchs', '', '', 'alterBranchName', 'T', document.frmAddBranch.BranchNameTrad);"><img src="images/trad.gif" alt="<%=getadminBranchsLngStr("DtxtTranslate")%>" border="0"></a></td>
				<td style="padding-left: 2px;" bgcolor="#F5FBFE">
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td><input type="checkbox" name="Active" value="Y" id="fp1" class="noborder"></td>
						<td><font face="Verdana" color="#4783C5" size="1"><label for="fp1"><%=getadminBranchsLngStr("DtxtActive")%></label></font></td>
					</tr>
				</table>
				</td>
			</tr>
			<tr>
				<td style="width: 100px" bgcolor="#E1F3FD"><font color="#31659C" face="Verdana" size="1">
				<strong><%=getadminBranchsLngStr("DtxtWarehouse")%></strong></font></td>
				<td style="width: 400px" bgcolor="#F5FBFE"><select size="1" name="WhsCode" class="input">
				<option></option>
				<% do while not rs.eof %>
				<option value="<%=rs("WhsCode")%>"><%=myHTMLEncode(rs("WhsName"))%></option>
				<% rs.movenext
				loop %>
				</select></td>
				<td bgcolor="#F5FBFE">&nbsp;</td>
				<td style="padding-left: 2px;" bgcolor="#F5FBFE">&nbsp;</td>
			</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		&nbsp;</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table7">
			<tr>
				<td width="77">
				<input type="submit" value="<%=getadminBranchsLngStr("DtxtAdd")%>" name="B1" class="OlkBtn"></td>
				<td><hr size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
		<input type="hidden" name="cmd" value="addBranch">
		<input type="hidden" name="redir" value="adminBranchs.asp">
	</form>
	<tr>
		<td>&nbsp;</td>
	</tr>
	<% If Not rd.Eof Then %>
	<form method="POST" action="submitBranchs.asp">
	<tr>
		<td bgcolor="#E1F3FD"><b><font size="1" color="#31659C" face="Verdana">
		<%=getadminBranchsLngStr("LttlEditBranch")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif">&nbsp;<font color="#4783C5" face="Verdana" size="1"><%=getadminBranchsLngStr("LttlEditBranchNote")%></font></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0">
			<tr>
				<td style="width: 16px" bgcolor="#E1F3FD"></td>
				<td style="width: 400px" bgcolor="#E1F3FD" class="style1"><font size="1" color="#31659C" face="Verdana">
				<strong><%=getadminBranchsLngStr("DtxtName")%></strong></font></td>
				<td style="width: 200px" bgcolor="#E1F3FD" class="style1"><font size="1" color="#31659C" face="Verdana">
				<strong><%=getadminBranchsLngStr("DtxtWarehouse")%></strong></font></td>
				<td bgcolor="#E1F3FD" class="style1"><font size="1" color="#31659C" face="Verdana">
				<strong><%=getadminBranchsLngStr("DtxtActive")%></strong></font></td>
				<td style="width: 16px" bgcolor="#E1F3FD"></td>
			</tr>
			<% do while not rd.eof %>
			<tr>
				<td style="width: 16px" bgcolor="#F3FBFE">
				<font size="1">
				<a href="adminBranchsEdit.asp?branchIndex=<%=rd("branchIndex")%>">
				<font color="#4783C5" face="Verdana">
				<img border="0" src="images/<%=Session("rtl")%>flechaselec.gif" width="15" height="13"></font></a></font></td>
				<td style="width: 400px" bgcolor="#F3FBFE">
				<table cellpadding="0" cellspacing="0" border="0" width="100%">
					<tr>
						<td>
						<input type="text" name="BranchName<%=rd("branchIndex")%>" size="24" style=" width: 100%" class="input" onkeydown="return chkMax(event, this, 50);" value="<% If Not IsNull(rd("branchName")) Then %><%=Server.HTMLEncode(rd("branchName"))%><% End If %>">
						</td>
						<td width="16"><a href="javascript:doFldTrad('Branchs', 'branchIndex', <%=rd("branchIndex")%>, 'alterBranchName', 'T', null);"><img src="images/trad.gif" alt="<%=getadminBranchsLngStr("DtxtTranslate")%>" border="0"></a></td>
					</tr>
				</table>
				</td>
				<td style="width: 200px" bgcolor="#F3FBFE">
				<font size="1" color="#4783C5" face="Verdana">
						<%=rd("whsName")%>&nbsp;
				</font></td>
						<td style="width: 100px" bgcolor="#F3FBFE" class="style1">
				<input <%If rd("Active") = "Y" then response.write "checked"%> type="checkbox" name="Active<%=rd("branchIndex")%>" value="Y" id="Active<%=rd("branchIndex")%>" class="noborder"></td>
						<td style="width: 16px" bgcolor="#F3FBFE">
				<a href="javascript:delBranch('<%=Replace(rd("branchName"),"'","\'")%>', <%=rd("branchIndex")%>);">
				<font color="#4783C5" face="Verdana" size="1">
				<img border="0" src="images/remove.gif" width="16" height="16"></font></a></td>
			</tr>
			<%  rd.movenext
			loop %>
		</table>
		</td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		&nbsp;</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table7">
			<tr>
				<td width="77">
				<input type="submit" value="<%=getadminBranchsLngStr("DtxtSave")%>" name="B1" class="OlkBtn"></td>
				<td><hr size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
		<input type="hidden" name="cmd" value="activeBranch">
		<input type="hidden" name="redir" value="adminBranchs.asp">
	</form>
	<tr>
		<td><font face="Verdana" size="1">&nbsp;
		</font>
		</td>
	</tr>
	<% End If %>
</table>
<script language="javascript">
function delBranch(branchName, branchIndex)
{
	if(confirm('<%=getadminBranchsLngStr("LtxtConfDelBranch")%>'.replace('{0}', branchName)))window.location.href='submitBranchs.asp?cmd=delBranch&branchIndex=' + branchIndex+ '&redir=adminBranchs.asp';
}

</script>
<!--#include file="adminTradSubmit.asp"-->
<!--#include file="bottom.asp" -->
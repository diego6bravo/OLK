<% pageTtl = "Reportes" %>
<!--#include file="repTop.inc" -->
<!--#include file="lang/adminReps.asp" -->
<!--#include file="adminTradSubmit.asp"-->

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style type="text/css">
.style1 {
	text-align: center;
}
</style>
</head>

<script language="javascript">
function Start(name, page, w, h, s, r) {
var winleft = (screen.width - w) / 2;
var winUp = (screen.height - h) / 2;
OpenWin = this.open(page, name, "toolbar=no,menubar=no,location=no,left="+winleft+",top="+winUp+",scrollbars="+s+",resizable="+r+", width="+w+",height="+h);
OpenWin.focus()
}
function doExpand(rgIndex)
{
	var sign = document.getElementById('signExpand' + rgIndex);
	
	var tr = document.getElementById('tr' + rgIndex);
	var display = false;

	if (tr.style.display == 'none')
	{
		tr.style.display = '';
		display = true;
	}
	else
	{
		tr.style.display = 'none';
		display = false;
	}
		
	sign.innerHTML = (display ? '[-]' : '[+]');
}
</script>
<% 

If repTbl = "OLK" Then
	uType = Request("uType")
Else
	uType = "V"
End If
sql = 	"select rsIndex, rsName + Case LinkOnly When 'N' Then N'' Else N' (" & getadminRepsLngStr("DtxtLink") & ")' end rsName, T0.rgIndex, rgName, " & _
		"Case SuperUser When 'Y' Then N'" & getadminRepsLngStr("LtxtSuperUser") & "' When 'N' Then N'" & getadminRepsLngStr("DtxtAll") & "' End Access, " & _
		"T0.Active, T0.rgIndex, LastUpdate UpdateDate, Convert(nvarchar(10),LastUpdate,108) UpdateTime " & _
		"from " & repTbl & "rs T0 " & _
		"inner join " & repTbl & "rg T1 on T1.rgIndex = T0.rgIndex " & _
		"where T1.UserType = '" & uType & "' " & _
		" order by 4, 2"
rs.open sql, conn, 3, 1

set rd = Server.CreateObject("ADODB.RecordSet")
sql = 	"select rgIndex, rgName, SuperUser, (select count('A') from " & repTbl & "RS where rgIndex = " & repTbl & "RG.rgIndex) rsCount " & _
		"from " & repTbl & "RG where UserType = '" & uType & "'"

If Request("uType") = "V" Then sql = sql & " and rgIndex >= 0 "
		
sql = sql & " order by rgName asc"
rd.open sql, conn, 3, 1
			%>
<table border="0" cellpadding="0" width="100%" id="table3">
	<tr class="TblRepTlt">
		<td>&nbsp;<%=getadminRepsLngStr("LttlRepLst")%></td>
	</tr>
	<form method="POST" action="repSubmit.asp" name="frmReps">
	<tr class="TblRepNrm">
		<td>
		<p align="justify"><img src="images/lentes.gif"> 
		<%=getadminRepsLngStr("LttlRepLstNote")%></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table12">
			<thead>
			<tr class="TblRepTlt">
				<td align="center" width="15"><b>
				<font size="1" face="Verdana" color="#31659C"></font></b></td>
				<td align="center">
				<%=getadminRepsLngStr("DtxtName")%></td>
				<% If uType = "V" Then %>
				<td align="center" style="width: 100px">
				<%=getadminRepsLngStr("DtxtAccess")%></td><% End If %>
				<td align="center" style="width: 140px">
				<%=getadminRepsLngStr("LtxtLastUpdate")%></td>
				<td align="center" style="width: 80px">
				<%=getadminRepsLngStr("DtxtActive")%></td>
				<td align="center" style="width: 16px"></td>
			</tr>
			</thead>
			<% 
			LastGroup = ""
			EndBody = False
			LastRep = False
			do While NOT RS.EOF
			If LastGroup <> rs("rgName") Then
			If EndBody Then 
				Response.Write "</tbody>"
			End If
			EndBody = True %>
			<thead>
			<tr class="TblRepTltNoBold" style="cursor: hand; " onclick="javascript:doExpand(<%=rs("rgIndex")%>);">
			  <td width="15" align="center">
			  	<span id="signExpand<%=rs("rgIndex")%>">[+]</span>
				</td>
				<td colspan="5"><%=rs("rgName")%></td>
				<% If uType = "V" Then %><% End If %>
				</tr>
			</thead>
			<% 
			LastGroup = rs("rgName")
			FirstRep = True
			End If
			If FirstRep Then %>
			<tbody id="tr<%=rs("rgIndex")%>" style="display: none; ">
			<% 
			FirstRep = False
			End If %>
			<tr class="TblRepTbl">
			  <td width="15">
				<a href="adminRepEdit.asp?rsIndex=<%=rs("rsIndex")%>"><img border="0" src="images/<%=Session("rtl")%>flechaselec.gif"></a></td>
				<td>
				<table cellpadding="0" cellspacing="0" border="0" width="100%">
					<tr class="TblRep<% If Alter Then %>A<% End If %>Tbl">
						<td><%=rs("rsName")%></td>
						<td align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
						<a href="javascript:doFldTrad('RS', 'rsIndex', '<%=rs("rsIndex")%>', 'alterRSName', 'T', null);"><img src="images/trad.gif" alt="<%=getadminRepsLngStr("DtxtTranslate")%>" border="0"></a>
						</td>
					</tr>
				</table>
				</td>
				<% If uType = "V" Then %><td style="width: 100px" class="style1"><%=rs("Access")%>&nbsp;</td><% End If %>
				<td style="width: 140px; text-align: center;">
				<%=FormatDate(rs("UpdateDate"), True)%>&nbsp;<%=rs("UpdateTime")%></td>
				<td style="width: 80px">
				<p align="center">
				<input class="OptionButton" type="checkbox" name="Active<%=rs("rsIndex")%>" <% If rs("Active") = "Y" Then %>checked<% End If %> value="Y"></td>
				<td style="width: 16px">
				<% If rs("rsIndex") >= 0 Then %>
				<a href="javascript:if(confirm('<%=getadminRepsLngStr("LtxtConfDelQry")%>'.replace('{0}', '<%=rs("rsName")%>')))window.location.href='repSubmit.asp?cmd=remRS&rsIndex=<%=rs("rsIndex")%>&UserType=<%=uType%>&rgIndex=<%=Request("rgIndex")%>'">
				<img border="0" src="images/remove.gif"></a><% Else %>&nbsp;<% End If %></td>
			</tr>
	<% 	RS.MoveNext
		loop
			If EndBody Then 
				Response.Write "</tbody>"
			End If
 %>
		  </table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table21">
			<tr>
				<td width="77">
				<input class="BtnRep" type="submit" value="<%=getadminRepsLngStr("DtxtSave")%>" name="btnSave"></td>
				<td width="77">
				<input class="BtnRep" type="button" value="<%=getadminRepsLngStr("DtxtNew")%>" name="B3" onclick="javascript:<% If rd.recordcount > 0 Then %>window.location.href='adminRepNew.asp?UserType=<%=uType%>'<% Else %>alert('<%=getadminRepsLngStr("LtxtValNoGrp")%>');<% End If %>"></td>
				<td><hr size="1"></td>
				<td width="77">
				<input class="BtnRep" type="button" value="<%=getadminRepsLngStr("LtxtImport")%>" name="btnImport" onclick="javascript:Start('rsExport', 'adminRepImport.asp?UserType=<%=uType%>&pop=Y', 600, 400, 'yes', 'no')"></td>
				<td width="77">
				<input class="BtnRep" type="button" value="<%=getadminRepsLngStr("LtxtExport")%>" name="btnExport" onclick="javascript:Start('rsExport', 'adminRepExport.asp?UserType=<%=uType%>&rgIndex=<%=Request("rgIndex")%>&pop=Y', 600, 400, 'yes', 'no')"></td>
			</tr>
		</table>
		</td>
	</tr>
		<input type="hidden" name="cmd" value="uActive">
		<input type="hidden" name="UserType" value="<%=uType%>">
	</form>
	<script language="javascript">
	function valFrm2()
	{
		chkRg = document.form2.chkrgName;
		{
		if (chkRg != null)
		{
			if (chkRg.length)
			{
				for (var i = 0;i<chkRg.length;i++)
				{
					if (chkRg[i].value == '')
					{
						alert('<%=getadminRepsLngStr("LtxtValGrpNam")%>');
						chkRg[i].focus();
						return false;					
					}
				}
			}
			else
				if (chkRg.value == '')
				{
					alert('<%=getadminRepsLngStr("LtxtValGrpNam")%>');
					chkRg.focus();
					return false;
				}
			}
		}
	}
	</script>
	<form method="POST" action="repSubmit.asp" name="form2" onsubmit="javascript:return valFrm2();">
	<tr class="TblRepTlt">
		<td>&nbsp;<%=getadminRepsLngStr("DtxtGroups")%></td>
	</tr>
	<tr class="TblRepNrm">
		<td>
		<p align="justify"><img src="images/lentes.gif">
		<%=getadminRepsLngStr("LttlGrpNote")%></td>
	</tr>
	<tr>
		<td >
		<table border="0" cellpadding="0" id="table20"><tr class="TblRepTltSub">
				<td align="center"><%=getadminRepsLngStr("DtxtGroup")%>&nbsp;</td>
				<td align="center">
				<% If uType = "V" Then %><%=getadminRepsLngStr("DtxtAccess")%><% ElseIf uType = "C" Then %><%=getadminRepsLngStr("DtxtActive")%><% End If %>&nbsp;</td>
				<td align="center" width="1"></td>
			</tr>
			<% do while not rd.eof %>
			<tr class="TblRepTbl">
				<td valign="bottom">
				<p align="center">
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td><input type="text" id="chkrgName" <% If rd("rgIndex") < 0 Then %>readonly class="inputDis"<% End If %> name="rgName<%=rd("rgIndex")%>" size="100" value="<%=Server.HTMLEncode(rd("rgName"))%>" onkeydown="return chkMax(event, this, 100);"></td>
						<td><a href="javascript:doFldTrad('RG', 'rgIndex', <%=rd("rgIndex")%>, 'alterRGName', 'T', null);"><img src="images/trad.gif" alt="<%=getadminRepsLngStr("DtxtTranslate")%>" border="0"></a></td>
					</tr>
				</table>
				</td>
				<td valign="top" align="center">
				<select size="1" name="SuperUser<%=rd("rgIndex")%>">
				<option <% If rd("SuperUser") = "N" Then %>selected<% End If %> value="N">
				<%=getadminRepsLngStr("DtxtAll")%></option>
				<option <% If rd("SuperUser") = "Y" Then %>selected<% End If %> value="Y">
				<% Select Case uType
					Case "V" %><%=getadminRepsLngStr("LtxtSuperUser")%><%
					Case "C" %><%=getadminRepsLngStr("DtxtUser")%><%
				End Select %></option>
				</select></td>
				<td valign="top" style="width: 15px">
				<% If rd("rsCount") = 0 Then %><a href="javascript:if(confirm('<%=getadminRepsLngStr("LtxtConfDelGrp")%>'.replace('{0}', '<%=Replace(rd("rgName"), "'", "\'")%>')))window.location.href='repSubmit.asp?cmd=remRG&delIndex=<%=rd("rgIndex")%>&rgIndex=<%=Request("rgIndex")%>&UserType=<%=uType%>'"><img border="0" src="images/remove.gif"></a><% End If %></td>
			</tr>
			<% rd.movenext
			loop %>
			<tr class="TblRep<% If Alter Then %>A<% End If %>Tbl">
				<td valign="top">
				<p align="center">
				<input type="hidden" name="rgNameTrad">
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td><input type="text" name="rgName" size="100" onkeydown="return chkMax(event, this, 100);"></td>
						<td><a href="javascript:doFldTrad('RG', 'rgIndex', '', 'alterRGName', 'T', document.form2.rgNameTrad);"><img src="images/trad.gif" alt="<%=getadminRepsLngStr("DtxtTranslate")%>" border="0"></a></td>
					</tr>
				</table>
				</td>
				<td valign="top" align="center">
				<select size="1" name="SuperUser">
				<option value="N"><%=getadminRepsLngStr("DtxtAll")%></option>
				<option value="Y"><% Select Case uType
					Case "V" %><%=getadminRepsLngStr("LtxtSuperUser")%><%
					Case "C" %><%=getadminRepsLngStr("DtxtUser")%><%
				End Select %></option>
				</select></td>
				<td valign="top" style="width: 15px">
				&nbsp;</td>
			</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table5">
			<tr>
				<td width="77">
				<input class="BtnRep" type="submit" value="<%=getadminRepsLngStr("DtxtSave")%>" name="B2"></td>
				<td><hr size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
	<input type="hidden" name="cmd" value="uGrp">
	<input type="hidden" name="UserType" value="<%=uType%>">
</table>
</form>

<form name="frmOrder" method="post" action="adminReps.asp">
<input type="hidden" name="uType" value="<%=Request("uType")%>">
<input type="hidden" name="order" value="">
</form>
<script language="javascript">
function doOrder(order) { frmOrder.order.value = order; frmOrder.submit(); }
</script>
<!--#include file="repBottom.inc" -->
<!--#include file="top.asp" -->
<!--#include file="lang/adminBN.asp" -->
<!--#include file="adminTradSubmit.asp"-->

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style type="text/css">
.style1 {
	color: #4783C5;
	font-family: Verdana;
	font-size: xx-small;
}
</style>
</head>
<%
conn.execute("use [" & Session("olkdb") & "]")

	GroupId = Request("GroupID")
	If Request("GroupID") <> "-1" and Request("GroupId") <> "" Then
		strGroupID = "AND T0.GroupID = " & GroupId  & " "	
	Else
		strGroupId = " "
	End If
	Order = Request("hdnOrder")
	If Order <> "" Then 
		Select Case Order
			Case "Grupo" : strOrder = "T0.GroupID"
			Case "FechaIni": strOrder = "StartDate"
			Case "FechaFin": strOrder = "EndDate"	
		End Select
	Else 
		strOrder = "Convert(nvarchar(100),BannerDesc)"
	End If

	strSQL = 	"SELECT T1.GroupName, BannerID, IsNull(Convert(nvarchar(100),BannerDesc), '#' + Convert(nvarchar(20),BannerID)) BannerDesc, Link, Picture, StartDate, EndDate, Status "& _
			 	"FROM OLKBN T0 " & _
			 	"inner join OLKBNGroups T1 on T1.GroupID = T0.GroupID "& _
			 	"WHERE Status in ('A', 'N') " & _
			 	strGroupID & _
			 	"ORDER BY " & strOrder 
	set rstBanner = conn.Execute(strSQL)
	
	strSQL = 	"select GroupId, GroupName, SizeX, SizeY, " & _
				"(select count('A') from olkBN where GroupID = T0.GroupId and Status <> 'D') as cant from OLKBNGroups T0 " & _
			 	"where groupID >= 0 "

	set rs = Server.CreateObject("ADODB.RecordSet")
	rs.open strSql, conn, 3, 1
%>
<script type="text/javascript">
function chkThis(fld, min, dec, oldVal)
{
	if (!IsNumeric(fld.value))
	{
		fld.value = oldVal.value;
		alert('<%=getadminBNLngStr("DtxtValNumVal")%>');
		fld.focus();
	}
	else if (parseFloat(fld.value) < parseFloat(min))
	{
		fld.value = min;
		alert("<%=getadminBNLngStr("DtxtValNumMinVal")%>".replace('{0}', min));
		fld.focus();
	}
	else if (parseFloat(fld.value) > 32727)
	{
		fld.value = 32727;
		alert("<%=getadminBNLngStr("DtxtValNumMaxVal")%>".replace('{0}', '32727'));
		fld.focus();
	}
	else
	{
		fld.value = parseInt(fld.value);
	}
	
	if (fld.value != '') fld.value = formatNumber(fld.value, dec).replace(',', '');
	oldVal.value = fld.value;
}
</script>
<table border="0" cellpadding="0" width="100%">
	<tr class="TblRepTlt">
		<td><%=getadminBNLngStr("LtxtBanList")%></td>
	</tr>
	<tr class="TblRepNrm">
		<td>
		<p align="justify"><img src="images/lentes.gif">
		<%=getadminBNLngStr("LtxtBanListDesc")%></td>
	</tr>
	<tr class="TblRepNrm">
		<td>
		<table border="0" cellpadding="0" width="100%" id="table12">
			<form method="POST" action="adminBN.asp" name="frmGroups">
				<tr class="TblRepTltSub">
					<td>
					<select size="1" name="GroupId" onchange="javascript:frmGroups.submit()">
					<option value="" <%if request("groupid") = "" then%>selected<%end if%>>
					<%=getadminBNLngStr("DtxtAll")%></option>
					<% If Not rs.Eof Then
					rs.Filter = "cant > 0"
					Do while Not rs.EOF %>
					<option value="<%=rs("GroupId")%>" <% If CStr(rs("groupid")) = Request("GroupId") Then %>selected<% End If %>>
					<%=rs("GroupName")%></option>
					<% rs.MoveNext
					Loop
					rs.Filter = ""
					rs.movefirst
					End If %>
					</select>
					<input type="hidden" name="order" value="<%=strOrder%>"></td>
				</tr>
			</form>
		</table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr class="TblRepTltSubNoUnder">
				<td align="center" style="width: 16px"></td>
				<% If Request("GroupId") = "" Then %>
				<td align="center"><%=getadminBNLngStr("DtxtGroup")%></td><% End If %>
				<td align="center"><%=getadminBNLngStr("DtxtDescription")%></td>
				<td align="center"><%=getadminBNLngStr("LtxtStartDate")%></td>
				<td align="center"><%=getadminBNLngStr("LtxtEndDate")%></td>
				<td align="center"><%=getadminBNLngStr("DtxtActive")%></td>
				<td align="center" style="width: 16px"></td>
			</tr>
			<form method="POST" action="adminSubmit.asp">
				<% 
				Do while not rstBanner.EOF %>
				<tr>
					<td style="width: 16px" bgcolor="#F3FBFE">
					<a href="adminBNEdit.asp?BannerID=<%=rstBanner("BannerID")%>&GroupId=<%=Request("GroupId")%>">
					<img border="0" src="images/<%=Session("rtl")%>flechaselec.gif"></a></td>
					<% If Request("GroupId") = "" Then %>
					<td class="style1" bgcolor="#F3FBFE"><%=rstBanner("GroupName")%></td><% End If %>
					<td bgcolor="#F3FBFE">
					<table cellpadding="0" cellspacing="0" border="0" width="100%">
						<tr>
							<td class="style1"><%=rstBanner("BannerDesc")%></td>
							<td align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><a href="javascript:doFldTrad('BN', 'BannerID', <%=rstBanner("BannerID")%>, 'AlterBannerDesc', 'T', null);"><img src="images/trad.gif" alt="<%=getadminBNLngStr("DtxtTranslate")%>" border="0"></a></td>
						</tr>
					</table>
					</td>
					<td align="center" class="style1" bgcolor="#F3FBFE"><%=FormatDate(rstBanner("StartDate"), True)%></td>
					<td align="center" class="style1" bgcolor="#F3FBFE"><%=FormatDate(rstBanner("EndDate"), True)%></td>
					<td align="center" bgcolor="#F3FBFE">
					<input type="checkbox" name="chkStatus<%=rstBanner("BannerID")%>" value="A" class="noborder" <% If rstbanner("status") = "A" then%> checked<% end if%>>
					</td>
					<td align="center" bgcolor="#F3FBFE" style="width: 16px">
					<a href="javascript:if(confirm('<%=getadminBNLngStr("LtxtConfRemBanner")%>'.replace('{0}', '<%=Replace(Server.HTMLEncode(rstBanner("BannerDesc")), "'", "\'")%>')))window.location.href='adminSubmit.asp?submitCmd=adminBN&cmd=remBN&BannerID=<%=rstBanner("BannerID")%>&GroupId=<%=Request("GroupID")%>'">
					<img border="0" src="images/remove.gif"></a></td>
				</tr>
				<%
				rstBanner.MoveNext
				Loop
				%>
				<tr>
					<td colspan="7">
					<input type="hidden" name="cmd" value="uActive">
					<input type="hidden" name="GroupId" value="<%=Request("GroupId")%>">
					<table border="0" cellspacing="0" width="100%" id="table13">
						<tr>
							<td width="77">
							<input type="submit" value="<%=getadminBNLngStr("DtxtSave")%>" name="btnGuardar" class="OlkBtn"></td>
							<td width="77">
							<input type="button" value="<%=getadminBNLngStr("DtxtNew")%>" name="btnNuevo" onclick="javascript:<% If rs.recordcount > 0 Then %>window.location.href='adminBNEdit.asp?groupId=<%=GroupId%>'<% Else %>alert('<%=getadminBNLngStr("LtxtValNoGrp")%>');<% End If %>" class="OlkBtn"></td>
							<td><hr size="1"></td>
						</tr>
					</table>
					</td>
				</tr>
				<input type="hidden" name="submitCmd" value="adminBN">
			</form>
		</table>
		</td>
	</tr>
	<tr class="TblRepTlt">
		<td><%=getadminBNLngStr("DtxtGroups")%></td>
	</tr>
	<tr class="TblRepNrm">
		<td>
		<p align="justify"><img src="images/lentes.gif">
		<%=getadminBNLngStr("LtxtGrpsDesc")%></td>
	</tr>
	<form method="POST" action="adminSubmit.asp" name="frmBNGrp" onsubmit="return valFrmBnGrp();">
	<tr>
		<td>
		<table border="0" cellpadding="0" id="table20">
			<script language="javascript">
			function valFrmBnGrp()
			{
				if (document.frmBNGrp.txtGroupName)
				{
					if (document.frmBNGrp.txtGroupName.length)
					{
						for (var i = 0;i<document.frmBNGrp.txtGroupName.length;i++)
						{
							if (document.frmBNGrp.txtGroupName[i].value == '')
							{
								alert('<%=getadminBNLngStr("LtxtValGrpNam")%>');
								document.frmBNGrp.txtGroupName[i].focus();
								return false;
							}
						}
					}
					else
					{
						if (document.frmBNGrp.txtGroupName.value == '')
						{
							alert('<%=getadminBNLngStr("LtxtValGrpNam")%>');
							document.frmBNGrp.txtGroupName.focus();
							return false;
						}
					}
				}
				if (document.frmBNGrp.NewGroupName.value != '' && (document.frmBNGrp.NewSizeX.value == '' || document.frmBNGrp.NewSizeY.value == ''))
				{
					alert('<%=getadminBNLngStr("LtxtValGrpSize")%>');
					if (document.frmBNGrp.NewSizeX.value == '') document.frmBNGrp.NewSizeX.focus();
					else document.frmBNGrp.NewSizeY.focus();
					return false;
				}
				return true;
			}
			</script>
			<input type="hidden" name="NewGroupNameTrad">
				<tr align="center" class="TblRepTltSubNoUnder">
					<td style="width: 300px"><%=getadminBNLngStr("DtxtGroup")%></td>
					<td><%=getadminBNLngStr("LtxtWidth")%></td>
					<td><%=getadminBNLngStr("LtxtHeight")%></td>
					<td><%=getadminBNLngStr("LtxtTag")%></td>
					<td style="width: 16px"></td>
				</tr>
				<% Alter = True
				Do While Not rs.EOF %>
				<tr>
					<td bgcolor="#F3FBFE" style="width: 300px">
					<table cellpadding="0" cellspacing="0" border="0" width="100%">
						<tr>
							<td><input type="text" style="width: 100%;" name="txtGroupName<%=rs("GroupId")%>" id="txtGroupName" size="30" maxlength="100" value="<%=Server.HTMLEncode(rs("GroupName"))%>" onkeydown="return chkMax(event, this, 100);"></td>
							<td style="width: 16px"><a href="javascript:doFldTrad('BNGroups', 'GroupID', <%=rs("GroupID")%>, 'alterGroupName', 'T', null);"><img src="images/trad.gif" alt="<%=getadminBNLngStr("DtxtTranslate")%>" border="0"></a></td>
						</tr>
					</table>
					</td>
					<td align="center" bgcolor="#F3FBFE">
					<input type="text" name="txtSizeX<%=rs("GroupId")%>" size="5" maxlength="100" value="<%=rs("SizeX")%>" style="text-align: right" onchange="javascript:chkThis(this, 1, 0, oldtxtSizeX<%=rs("GroupId")%>);">
					<input type="hidden" id="oldtxtSizeX<%=rs("GroupId")%>" value="<%=rs("SizeX")%>"></td>
					<td align="center" bgcolor="#F3FBFE">
					<input type="text" name="txtSizeY<%=rs("GroupId")%>" size="5" maxlength="100" value="<%=rs("SizeY")%>" style="text-align: right" onchange="javascript:chkThis(this, 1, 0, oldtxtSizeY<%=rs("GroupId")%>);">
					<input type="hidden" id="oldtxtSizeY<%=rs("GroupId")%>"value="<%=rs("SizeY")%>"></td>
					<td align="center" class="style1" bgcolor="#F3FBFE">
					&lt;!--doBanner<%=rs("GroupId")%>--&gt;</td>
					<td align="center" bgcolor="#F3FBFE" style="width: 16px"><% If rs("cant") = 0 then%><a href="javascript:if(confirm('<%=getadminBNLngStr("LtxtConfDelGrp")%>'.replace('{0}', '<%=Replace(rs("GroupName"), "'", "\'")%>')))window.location.href='adminSubmit.asp?submitCmd=adminBN&cmd=remGR&delId=<%=rs("GroupId")%>'"><img border="0" src="images/remove.gif"></a><%Else%>
					<% End If%></td>
				</tr>
				<% Alter = Not Alter
				rs.MoveNext
				loop %>
				<tr>
					<td bgcolor="#F3FBFE" style="width: 300px">
					<table cellpadding="0" cellspacing="0" border="0" width="100%">
						<tr>
							<td>
							<input type="text" style="width: 100%;" name="NewGroupName" size="30" maxlength="100" onkeydown="return chkMax(event, this, 100);">
							</td>
							<td style="width: 16px"><a href="javascript:doFldTrad('BNGroups', 'GroupID', '', 'alterGroupName', 'T', document.frmBNGrp.NewGroupNameTrad);"><img src="images/trad.gif" alt="<%=getadminBNLngStr("DtxtTranslate")%>" border="0"></a>
							</td>
						</tr>
					</table></td>
					<td align="center" bgcolor="#F3FBFE">
					<input type="text" name="NewSizeX" size="5" maxlength="100" style="text-align: right" onchange="javascript:if (this.value != '')chkThis(this, 1, 0, oldNewSizeX);">
					<input type="hidden" id="oldNewSizeX" value=""></td>
					<td align="center" bgcolor="#F3FBFE">
					<input type="text" name="NewSizeY" size="5" maxlength="100" style="text-align: right" onchange="javascript:if (this.value != '')chkThis(this, 1, 0, oldNewSizeY);">
					<input type="hidden" id="oldNewSizeY" value=""></td>
					<td align="center" bgcolor="#F3FBFE">
					&nbsp;</td>
					<td align="center" bgcolor="#F3FBFE" style="width: 16px">&nbsp;</td>
				</tr>
		</table>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td>
		<input type="hidden" name="Cmd" value="uGrp">
		<input type="hidden" name="GroupId" value="<%=GroupId%>">
		<table cellpadding="0" cellspacing="0" border="0" width="100%">
			<tr>
				<td width="77"><input type="submit" value="<%=getadminBNLngStr("DtxtSave")%>" name="btnSave" class="OlkBtn"></td>
				<td><hr size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
	<input type="hidden" name="submitCmd" value="adminBN">
</form>
	<form method="POST" action="adminBN.asp" name="frmOrder">
		<tr>
			<td><input type="hidden" name="hdnOrder" size="20" value></td>
		</tr>
	</form>
</table>
<script language="javascript">
	function doOrder(order) {
		frmOrder.hdnOrder.value = order; 
		frmOrder.submit(); 
	}
</script>
<!--#include file="bottom.asp" -->
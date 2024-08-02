<!--#include file="top.asp" -->
<!--#include file="lang/adminMenuGroups.asp" -->
<!--#include file="adminTradSubmit.asp"-->
<% conn.execute("use [" & Session("olkdb") & "]")
set rd = Server.CreateObject("ADODB.RecordSet") %>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style type="text/css">
.style1 {
	background-color: #E2F3FC;
}
.style2 {
	font-weight: bold;
	background-color: #E2F3FC;
}
.style3 {
	background-color: #F3FBFE;
}
.style6 {
	background-color: #E2F3FC;
	font-family: Verdana;
	font-size: xx-small;
	color: #31659C;
}
.style7 {
	background-color: #F3FBFE;
	font-family: Verdana;
	font-size: xx-small;
	color: #31659C;
}
.style8 {
	background-color: #F3FBFE;
}
.style12 {
	background-color: #E2F3FC;
	font-size: xx-small;
	color: #31659C;
}
.style13 {
	background-color: #E2F3FC;
	font-family: Verdana;
	font-size: xx-small;
	color: #31659C;
	text-align: center;
	font-weight: bold;
}
</style>
</head>
<script type="text/javascript">
function delGroup(GroupID, GroupName)
{
	if(confirm('<%=getadminMenuGroupsLngStr("LtxtConfRemTree")%>'.replace('{0}', GroupName)))
		window.location.href='adminSubmit.asp?submitCmd=menuGroups&cmd=remGroup&GroupID=' + GroupID;
}
function delLines(LineID, LineName)
{
	if (confirm('<%=getadminMenuGroupsLngStr("LtxtConfRemLine")%>'.replace('{0}', LineName)))
		window.location.href='adminSubmit.asp?submitCmd=menuGroups&cmd=remLine&GroupID=<%=Request("editID")%>&LineID=' + LineID;
}

function valFrm()
{
	varGroupName = document.frmMenuGroups.GroupName;
	if (varGroupName.length)
	{
		for (var i = 0;i<varGroupName.length;i++)
		{
			if (varGroupName[i].value == '')
			{
				alert('<%=getadminMenuGroupsLngStr("LtxtValGrpNam")%>');
				varGroupName[i].focus();
				return false;
			}
		}
	}
	else
	{
		if (varGroupName.value == '')
		{
			alert('<%=getadminMenuGroupsLngStr("LtxtValGrpNam")%>');
			varGroupName.focus();
			return false;
		}
	}
	
	var hasDefault = false;
	for (var i = 0;i<document.frmMenuGroups.chkActive.length;i++)
	{
		if (document.frmMenuGroups.rdDefault[i].checked)
		{
			if (document.frmMenuGroups.chkActive[i].checked) hasDefault = true;
			else
			{
				alert('<%=getadminMenuGroupsLngStr("LtxtValActiveDefTree")%>');
				return false;
			}
		}
	}
	if (!hasDefault)
	{
		alert('<%=getadminMenuGroupsLngStr("LtxtValDefSearchTree")%>');
		return false;
	}
	
	return true;
}
</script>
<script language="javascript" src="js_up_down.js"></script>
<% If Request("new") <> "Y" and Request("editID") = "" Then %>
<table border="0" cellpadding="0" width="100%">
	<form name="frmMenuGroups" method="post" action="adminSubmit.asp" onsubmit="javascript:return valFrm();">
	<tr>
		<td bgcolor="#E1F3FD"><b><font color="#FFFFFF">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><%=getadminMenuGroupsLngStr("LttlSearchTree")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1"> 
		</font><font face="Verdana" size="1" color="#4783C5">
		<%=getadminMenuGroupsLngStr("LttlSearchTreeDesc")%></font></td>
	</tr>
	<tr>
		<td align="center">
		<table border="0" cellpadding="0" width="100%">
			<tr>
				<td align="center" width="1" class="style2">&nbsp;</td>
				<td align="center" class="style6">
				<strong><%=getadminMenuGroupsLngStr("DtxtName")%></strong></td>
				<td align="center" class="style6">
				<strong><%=getadminMenuGroupsLngStr("DtxtDefault2")%></strong></td>
				<td align="center" class="style6">
				<strong><%=getadminMenuGroupsLngStr("DtxtPosition2")%></strong></td>
				<td align="center" class="style6">
				<strong><%=getadminMenuGroupsLngStr("DtxtActive")%></strong></td>
				<td align="center" class="style6">
				<strong><%=getadminMenuGroupsLngStr("DtxtOrder")%></strong></td>
				<td align="center" class="style1" style="width: 16px"></td>
			</tr>
			<% 
			sql = 	"select T0.GroupID, T0.GroupName, T0.AllPosition, T0.Active, T0.Ordr, " & _
					"case when (select DefMenuGroup from OLKCommon) = T0.GroupID Then 'Y' Else 'N' End [Default] " & _
					"from OLKMenuGroups T0 " & _
					"order by T0.Ordr "
			set rs = conn.execute(sql)
			do While NOT RS.EOF
			If rs("GroupID") >= 0 Then
				Name = rs("GroupName")
			Else
				If rs("GroupID") = -1 Then
					Name = getadminMenuGroupsLngStr("DtxtGroup") & " / " & getadminMenuGroupsLngStr("DtxtFirm")
				ElseIf rs("GroupID") = -2 Then
					Name = getadminMenuGroupsLngStr("DtxtFirm") & " / " & getadminMenuGroupsLngStr("DtxtGroup")
				End If
			End If
			FieldID = Replace(rs("GroupID"), "-", "_") %>
			<tr class="TblRepTbl">
			  <td width="1" class="style3">
				<% If rs("GroupID") >= 0 Then %><a href="adminMenuGroups.asp?editID=<%=rs("GroupID")%>"><img border="0" src="images/<%=Session("rtl")%>flechaselec.gif"></a><% Else %>&nbsp;<% End If %></td>
				<td class="style3">
				<% If rs("GroupID") >= 0 Then %>
				<table cellpadding="0" border="0">
					<tr>
						<td><input type="text" id="GroupName" name="GroupName<%=FieldID%>" size="43" value="<%=myHTMLEncode(rs("GroupName"))%>" onkeydown="return chkMax(event, this, 100);"></td>
						<td><a href="javascript:doFldTrad('MenuGroups', 'GroupID', <%=rs("GroupID")%>, 'alterGroupName', 'T', null);"><img src="images/trad.gif" alt="<%=getadminMenuGroupsLngStr("DtxtTranslate")%>" border="0"></a></td>
					</tr>
				</table>
				<% Else %><%=Name%>
				<% End If %>
				</td>
				<td align="center" class="style3">
				<input name="rdDefault" type="radio" <% If rs("Default") = "Y" Then %>checked<% End If %> value="<%=rs("GroupID")%>" class="noborder">
				</td>
				<td align="center" class="style3">
				<select name="AllPosition<%=FieldID%>">
				<option value="B"><%=getadminMenuGroupsLngStr("DtxtDown")%></option>
				<option <% If rs("AllPosition") = "T" Then %>selected<% End If %> value="T"><%=getadminMenuGroupsLngStr("DtxtUp")%></option>
				</select></td>
				<td align="center" class="style3">
				<input type="checkbox" id="chkActive" name="chkActive<%=FieldID%>" <% If rs("Active") = "Y" Then %>checked<% End If %> value="Y" class="noborder">
				</td>
				<td align="center" class="style3">
				<table cellpadding="0" cellspacing="0" border="0" align="center">
					<tr>
						<td>
							<input type="text" name="GroupOrder<%=FieldID%>" id="GroupOrder<%=FieldID%>" size="7" style="text-align:right" class="input" onfocus="this.select();" onmouseup="event.preventDefault()" onkeydown="return chkMax(event, this, 6);" value="<%=rs("Ordr")%>">
						</td>
						<td valign="middle">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td><img src="images/img_nud_up.gif" id="btnGroupOrder<%=FieldID%>Up"></td>
							</tr>
							<tr>
								<td><img src="images/spacer.gif"></td>
							</tr>
							<tr>
								<td><img src="images/img_nud_down.gif" id="btnGroupOrder<%=FieldID%>Down"></td>
							</tr>
						</table></td>
					</tr>
				</table>
				</td>
				<td class="style3" style="width: 16px">
				<% If rs("GroupID") >= 0 Then %>
				<a href="javascript:delGroup(<%=rs("GroupID")%>, document.frmMenuGroups.GroupName<%=FieldID%>.value);">
				<img border="0" src="images/remove.gif"></a><% Else %>&nbsp;<% End If %></td>
			</tr>
			<script language="javascript">NumUDAttach('frmMenuGroups', 'GroupOrder<%=FieldID%>', 'btnGroupOrder<%=FieldID%>Up', 'btnGroupOrder<%=FieldID%>Down');</script>
		<% 	RS.MoveNext
			loop %>
		  </table>
		</td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE" align="center">
		<table border="0" cellpadding="0" width="100%" id="table7">
			<tr>
				<td width="77">
				<input type="submit" value="<%=getadminMenuGroupsLngStr("DtxtSave")%>" name="B1" class="OlkBtn"></td>
				<td width="77">
				<input type="button" value="<%=getadminMenuGroupsLngStr("DtxtNew")%>" name="B1" class="OlkBtn" onclick="javascript:window.location.href='adminMenuGroups.asp?new=Y'"></td>
				<td><hr size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
	<input type="hidden" name="cmd" value="update">
	<input type="hidden" name="submitCmd" value="menuGroups">
	</form>
</table>
<% Else 
If Request("editID") <> "" Then
	sql = "select GroupName, Active, Ordr, SearchFilter, AllPosition from OLKMenuGroups where GroupID = " & Request("editID") 
	set rs = conn.execute(sql)
	GroupName = rs("GroupName")
	Active = rs("Active")
	SearchFilter = rs("SearchFilter")
	AllPosition = rs("AllPosition")
Else
	sql = "select IsNull(Max(Ordr)+1, 1) Ordr from OLKMenuGroups"
	set rs = conn.execute(sql)
	SearchFilter = ""
End If

set rt = Server.CreateObject("ADODB.RecordSet")
sql = "select name from sysobjects where xtype in ('V', 'U') order by 1"
set rt = conn.execute(sql)

If Not rt.Eof and Request("editID") <> "" Then
	set rtCols = Server.CreateObject("ADODB.RecordSet")
	sql = "select T0.Name TableName, T1.Name ColName " & _
	"from sysobjects T0 " & _
	"inner join syscolumns T1 on T1.ID = T0.ID " & _
	"where T0.Name in (select DescTable from OLKMenuGroupsLines where GroupID = " & Request("editID") & " and DescTable is not null) "
	rtCols.open sql, conn, 3, 1
End If %>
<script type="text/javascript">
<!--
function valFrm2()
{
	if (document.frmEditMenuGroup.GroupName.value == '')
	{
		alert('<%=getadminMenuGroupsLngStr("LtxtValGrpNam")%>');
		document.frmEditMenuGroup.GroupName.focus();
		return false;
	}
	if (document.frmEditMenuGroup.valSearchFilter.value == 'Y')
	{
		alert('<%=getadminMenuGroupsLngStr("LtxtValCatFltQryVal")%>');
		document.frmEditMenuGroup.btnVerfyFilter.focus();
		return false;
	}
	return true;
}
function VerfyFilter()
{
	document.frmVerfyQuery.Query.value = document.frmEditMenuGroup.SearchFilter.value;
	if (document.frmVerfyQuery.Query.value != '')
	{
		document.frmVerfyQuery.submit();
	}
	else
	{
		VerfyQueryVerified();
	}
}
function VerfyQueryVerified()
{
	//document.frmEditMenuGroup.btnVerfyFilter.disabled = true;
	document.frmEditMenuGroup.btnVerfyFilter.src='images/btnValidateDis.gif'
	document.frmEditMenuGroup.btnVerfyFilter.style.cursor = '';
	document.frmEditMenuGroup.valSearchFilter.value='N';

}

//-->
</script>
<form name="frmVerfyQuery" action="verfyQuery.asp" method="post" target="iVerfyQuery">
	<iframe id="iVerfyQuery" name="iVerfyQuery" style="display: none" src=""></iframe>
	<input type="hidden" name="type" value="SearchTreeFilter">
	<input type="hidden" name="Query" value="">
	<input type="hidden" name="parent" value="Y">
</form>
<table border="0" cellpadding="0" width="100%">
	<form name="frmEditMenuGroup" method="post" action="adminSubmit.asp" onsubmit="javascript:return valFrm2();">
	<tr>
		<td bgcolor="#E1F3FD"><b><font color="#FFFFFF">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><% If Request("editID") <> "" Then %><%=getadminMenuGroupsLngStr("DtxtEdit")%><% ElseIf Request("new") = "Y" Then %><%=getadminMenuGroupsLngStr("DtxtAdd")%><% End If %>&nbsp;<%=getadminMenuGroupsLngStr("LttlSearchTree")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1"> 
		</font><font face="Verdana" size="1" color="#4783C5">
		<%=getadminMenuGroupsLngStr("LttlSearchTreeDesc")%></font></td>
	</tr>
	<tr>
		<td align="center">
		<table align="<% If Session("rtl") = "" Then %>left<% Else %>right<% End If %>" style="width: 100%">
			<tr>
				<td class="style6"><strong><%=getadminMenuGroupsLngStr("DtxtName")%></strong></td>
				<td class="style3">
				<table cellpadding="0" cellspacing="0" border="0" width="100%">
					<tr>
						<td><input type="text" style="width: 100%;" id="GroupName" name="GroupName" size="50" value="<%=myHTMLEncode(GroupName)%>" onkeydown="return chkMax(event, this, 100);"></td>
						<td style="width: 16px"><a href="javascript:doFldTrad('MenuGroups', 'GroupID', '<%=Request("editID")%>', 'alterGroupName', 'T', <% If Request("editID") = "" Then %>frmEditMenuGroup.GroupNameTrad<% Else %>null<% End If %>);"><img src="images/trad.gif" alt="<%=getadminMenuGroupsLngStr("DtxtTranslate")%>" border="0"></a></td>
					</tr>
				</table>
				</td>
			</tr>
			<tr>
				<td class="style6" style="width: 100px"><strong><%=getadminMenuGroupsLngStr("DtxtPosition2")%></strong></td>
				<td class="style3">
				<select name="AllPosition">
				<option value="B"><%=getadminMenuGroupsLngStr("DtxtDown")%></option>
				<option <% If AllPosition = "T" Then %>selected<% End If %> value="T"><%=getadminMenuGroupsLngStr("DtxtUp")%></option>
				</select>&nbsp;</td>
			</tr>
			<tr>
				<td class="style6" style="width: 100px"><strong><%=getadminMenuGroupsLngStr("DtxtOrder")%></strong></td>
				<td class="style3">
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td>
							<input type="text" name="GroupOrder" id="GroupOrder" size="7" style="text-align:right" class="input" onfocus="this.select();" onmouseup="event.preventDefault()" onkeydown="return chkMax(event, this, 6);" value="<%=rs("Ordr")%>">
						</td>
						<td valign="middle">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td><img src="images/img_nud_up.gif" id="btnGroupOrderUp"></td>
							</tr>
							<tr>
								<td><img src="images/spacer.gif"></td>
							</tr>
							<tr>
								<td><img src="images/img_nud_down.gif" id="btnGroupOrderDown"></td>
							</tr>
						</table></td>
					</tr>
				</table>
				</td>
			</tr>
			<tr>
				<td class="style6"></td>
				<td class="style3" align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
 				<table cellpadding="0" cellspacing="0" border="0">
 					<tr>
 						<td><input type="checkbox" name="chkActive" class="noborder" <% If Active = "Y" Then %>checked<% End If %> id="chkActive" value="Y"></td>
 						<td><label for="chkActive"><font size="1" color="#4783C5" face="Verdana"><strong><%=getadminMenuGroupsLngStr("DtxtActive")%></strong></font></label></td>
 					</tr>
 				</table>
				</td>
			</tr>
			<tr>
				<td valign="top" colspan="2" class="style6"><%=getadminMenuGroupsLngStr("LtxtFilter")%> - (ItemCode not in)</td>
			</tr>
			<tr>
				<td valign="top" class="style1"></td>
				<td class="style3">
				<table cellpadding="0" cellspacing="0" border="0" width="100%">
					<tr>
						<td rowspan="2">
							<textarea name="SearchFilter" rows="10" dir="ltr" cols="87" class="input" onkeydown="javscript:document.frmEditMenuGroup.btnVerfyFilter.src='images/btnValidate.gif';document.frmEditMenuGroup.btnVerfyFilter.style.cursor = 'hand';;document.frmEditMenuGroup.valSearchFilter.value='Y';" style="width: 100%;"><%=myHTMLEncode(searchFilter)%></textarea>
						</td>
						<td valign="top" width="1">
							<img style="cursor: hand; " src="images/qry_note.gif" id="btnNoteFilter" alt="<%=getadminMenuGroupsLngStr("DtxtDefinition")%>" onclick="javascript:doFldNote(15, 'SearchFilter', '<%=Request("editID")%>', <% If Request("editID") <> "" Then %>null<% Else %>document.frmEditMenuGroup.SearchFilterDef<% End If %>);">
						</td>
					</tr>
					<tr>
						<td valign="bottom" width="1">
							<img src="images/btnValidateDis.gif" id="btnVerfyFilter" alt="<%=getadminMenuGroupsLngStr("DtxtValidate")%>" onclick="javascript:if (document.frmEditMenuGroup.valSearchFilter.value == 'Y')VerfyFilter();">
							<input type="hidden" name="valSearchFilter" id="valSearchFilter" value="N">
						</td>
					</tr>
				</table>
				</td>
			</tr>
			<tr>
				<td style="width: 160px" align="<% If Session("rtl") = "" Then %>left<% Else %>right<% End If %>" valign="top" class="style12">
				<font face="Verdana"><strong><%=getadminMenuGroupsLngStr("DtxtVariables")%></strong></font></td>
				<td class="style3">
				<font size="1" color="#4783C5" face="Verdana"><span dir="ltr">@CardCode</span> = <%=getadminMenuGroupsLngStr("LtxtCCode")%></font></td>
			</tr>
		</table>
		<script language="javascript">NumUDAttach('frmEditMenuGroup', 'GroupOrder', 'btnGroupOrderUp', 'btnGroupOrderDown');</script>
		</td>
	</tr>
	<tr>
		<td align="center">
		<table border="0" cellpadding="0" width="100%">
			<tr>
				<td align="center" class="style6">
				<strong><%=getadminMenuGroupsLngStr("DtxtDescription")%></strong></td>
				<td align="center" class="style6">
				<strong><%=getadminMenuGroupsLngStr("DtxtType")%></strong></td>
				<td align="center" class="style6">
				<strong><%=getadminMenuGroupsLngStr("LtxtTable")%></strong></td>
				<td align="center" class="style6">
				<strong><%=getadminMenuGroupsLngStr("LtxtID")%></strong></td>
				<td align="center" class="style6">
				<strong><%=getadminMenuGroupsLngStr("LtxtNameID")%></strong></td>
				<td align="center" class="style6">
				<strong><%=getadminMenuGroupsLngStr("LtxtDescTable")%></strong></stong></td>
				<td align="center" class="style6">
				<strong><%=getadminMenuGroupsLngStr("LtxtDescID")%></strong></td>
				<td align="center" class="style6">
				<strong><%=getadminMenuGroupsLngStr("LtxtDescName")%></strong></td>
				<td align="center" class="style6">
				<strong><%=getadminMenuGroupsLngStr("DtxtOrder")%></strong></td>
				<td align="center" width="1" class="style1"></td>
			</tr>
			<% 
			If Request("editID") <> "" Then
			valDesc = ""
			sql = 	"select T0.LineID, " & _
					"Case T0.TableType " & _
					"	When 'S' Then " & _
					"		Case T0.QueryTable " & _
					"			When 'OITB' Then N'" & getadminMenuGroupsLngStr("DtxtGroup") & "' " & _
					"			When 'OMRC' Then N'" & getadminMenuGroupsLngStr("DtxtFirm") & "' " & _
					"			When 'OCRD' Then N'" & getadminMenuGroupsLngStr("LtxtSupplier") & "' " & _
					"		End " & _
					"	When 'U' Then T1.AliasID + ' - ' + T1.Descr " & _
					"	When 'F' Then T0.FilterID " & _
					"	When 'Q' Then N'" & getadminMenuGroupsLngStr("LtxtProp") & "' " & _
					"	When 'T' Then N'" & getadminMenuGroupsLngStr("LtxtTable") & "' " & _
					"End Name, T0.TableType, " & _
					"Case T0.TableType When 'S' Then N'" & getadminMenuGroupsLngStr("DtxtSystem") & "' When 'U' Then N'" & getadminMenuGroupsLngStr("DtxtUDF") & "' When 'F' Then N'" & getadminMenuGroupsLngStr("DtxtField") & "' When 'Q' Then N'" & getadminMenuGroupsLngStr("LtxtProp") & "' When 'T' Then N'" & getadminMenuGroupsLngStr("LtxtTable") & "' End TableTypeDesc, " & _
					"IsNull(T0.QueryTable, 'OITM') QueryTable, T0.FilterID, T0.FilterName, T0.FilterFormula, T0.DescTable, T0.DescID, T0.DescName, T0.DescFormula, T0.Ordr " & _
					"from OLKMenuGroupsLines T0 " & _
					"left outer join CUFD T1 on T1.TableID = 'OITM' and T0.TableType = 'U' and T1.AliasID = T0.FilterID " & _
					"where T0.GroupID = " & Request("editID") & " order by T0.Ordr asc"
			set rs = conn.execute(sql)
			If Not rs.Eof Then
			do While NOT RS.EOF
			FieldID = rs("LineID")
			If rs("TableType") = "F" Then
				If valDesc <> "" Then valDesc = valDesc & ", "
				valDesc = valDesc & FieldID
			End If %>
			<tr class="TblRepTbl">
				<td class="style3"><%=rs("Name")%>
				</td>
				<td align="center" class="style3">
				<%=rs("TableTypeDesc")%>&nbsp;</td>
				<td align="center" class="style3">
				<%=rs("QueryTable")%>&nbsp;</td>
				<td align="center" class="style3">
				<%=rs("FilterID")%><% Select Case rs("TableType") 
					Case "F" %>&nbsp;<input type="button" name="btnFormula" value="..." onclick="setFormula(<%=FieldID%>, 'F');">
					<% Case "Q"
					sql = "select ItmsTypCod from OLKMenuGroupsLinesQryGroups where GroupID = " & Request("editID") & " and LineID = " & rs("LineID")
					set rd = conn.execute(sql)
					strQryGroups = ""
					do while not rd.eof
						If strQryGroups <> "" Then strQryGroups = strQryGroups & ", "
						strQryGroups = strQryGroups & rd(0)
					rd.movenext
					loop
					 %>
					<input type="button" name="btnQryGroups<%=FieldID%>" id="btnQryGroups<%=FieldID%>" value="..." style="width: 120px;" onclick="getQryGroups('<%=FieldID%>')">
					<input type="hidden" name="QryGroups<%=FieldID%>" id="QryGroupsAdd" value="<%=strQryGroups%>">
				<% End Select %></td>
				<td align="center" class="style3">
				<%=rs("FilterName")%>&nbsp;</td>
				<td align="center" class="style3">
				<% If rs("TableType") = "F" Then %>
				<select size="1" name="cmbDescTable<%=FieldID%>" id="cmbDescTable<%=FieldID%>" onchange="javascript:changeDescTable(document.frmEditMenuGroup.cmbDescID<%=FieldID%>, document.frmEditMenuGroup.cmbDescName<%=FieldID%>, this.value);">
				<option></option>
				<%
				rt.movefirst
				do while not rt.eof %>
				<option <% If rt(0) = rs("DescTable") Then %>selected<% End If %> value="<%=rt(0)%>"><%=rt(0)%></option>
				<% rt.movenext
				loop %>
				</select><% End If %>&nbsp;</td>
				<td align="center" class="style3">
				<% If rs("TableType") = "F" Then
				rtCols.Filter = "TableName = '" & rs("DescTable") & "'" %>
				<select size="1" name="cmbDescID<%=FieldID%>">
				<option value="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</option>
				<% If Not IsNull(rs("DescTable")) Then
				do while not rtCols.eof %>
				<option <% If rtCols("ColName") = rs("DescID") Then %>selected<% End If %> value="<%=rtCols("ColName")%>"><%=rtCols("ColName")%></option>
				<% rtCols.movenext
				loop 
				rtCols.movefirst
				End If %>
				</select><% End If %>
				&nbsp;<% If rs("TableType") = "F" Then %><input type="button" name="btnFormula" value="..." onclick="setFormula(<%=FieldID%>, 'D');"><% End If %></td>
				<td align="center" class="style3">
				<% If rs("TableType") = "F" Then %>
				<select size="1" name="cmbDescName<%=FieldID%>">
				<option value="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</option>
				<% If Not IsNull(rs("DescTable")) Then
				do while not rtCols.eof %>
				<option <% If rtCols("ColName") = rs("DescName") Then %>selected<% End If %> value="<%=rtCols("ColName")%>"><%=rtCols("ColName")%></option>
				<% rtCols.movenext
				loop 
				rtCols.movefirst
				End If %>
				</select><% End If %>
				&nbsp;</td>
				<td align="center" class="style3">
				<table cellpadding="0" cellspacing="0" border="0" align="center">
					<tr>
						<td>
							<input type="text" name="LineOrder<%=FieldID%>" size="7" style="text-align:right" class="input" onfocus="this.select();" onmouseup="event.preventDefault()" onkeydown="return chkMax(event, this, 6);" value="<%=rs("Ordr")%>">
						</td>
						<td valign="middle">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td><img src="images/img_nud_up.gif" id="btnLineOrder<%=FieldID%>Up"></td>
							</tr>
							<tr>
								<td><img src="images/spacer.gif"></td>
							</tr>
							<tr>
								<td><img src="images/img_nud_down.gif" id="btnLineOrder<%=FieldID%>Down"></td>
							</tr>
						</table></td>
					</tr>
				</table>
				</td>
				<td width="1" class="style3">
				<a href="javascript:delLines(<%=rs("LineID")%>, '<%=rs("Name")%>');">
				<img border="0" src="images/remove.gif"></a></td>
			</tr>
			<input type="hidden" name="FilterFormula<%=FieldID%>" id="FilterFormula<%=FieldID%>" value="<% If Not IsNull(rs("FilterFormula")) Then %><%=Server.HTMLEncode(rs("FilterFormula"))%><% End If %>">
			<input type="hidden" name="DescFormula<%=FieldID%>" id="DescFormula<%=FieldID%>" value="<% If Not IsNull(rs("DescFormula")) Then %><%=Server.HTMLEncode(rs("DescFormula"))%><% End If %>">
			<script language="javascript">NumUDAttach('frmEditMenuGroup', 'LineOrder<%=FieldID%>', 'btnLineOrder<%=FieldID%>Up', 'btnLineOrder<%=FieldID%>Down');</script>
		<% 	NewOrdr = rs("Ordr")+1
			RS.MoveNext
			loop
			Else
			NewOrdr = 1 %>
			<tr class="TblRepTbl">
				<td colspan="10" align="center" class="style7"><strong><%=getadminMenuGroupsLngStr("DtxtNoData")%></strong></td>
			</tr>
		<%	End If
			Else %>
			<tr class="TblRepTbl">
				<td colspan="10" align="center" class="style7"><strong><%=getadminMenuGroupsLngStr("LtxtAddLine")%></strong></td>
			</tr>
		<%	End If %>
		<% If Request("editID") <> "" Then %>
			<tr>
				<td colspan="10">
				<table style="width: 100%">
					<tr>
						<td bgcolor="#E1F3FD" colspan="5"><b><font face="Verdana" size="1" color="#31659C"><%=getadminMenuGroupsLngStr("DtxtAdd")%>&nbsp;<%=getadminMenuGroupsLngStr("DtxtLine")%></font></b></td>
					</tr>
				</table>
				<table cellpadding="0" border="0">
					<tr>
						<td class="style13"><%=getadminMenuGroupsLngStr("DtxtType")%></td>
						<td class="style13"><span id="txtTblFld"><%=getadminMenuGroupsLngStr("LtxtTable")%></span></td>
						<td class="style13" id="tdField1" style="display: none; "><%=getadminMenuGroupsLngStr("DtxtField")%></td>
						<td class="style13" id="tdDescTable1" style="display: none; "><%=getadminMenuGroupsLngStr("LtxtDescTable")%></td>
						<td class="style13" id="tdDescTable2" style="display: none; "><%=getadminMenuGroupsLngStr("LtxtDescID")%></td>
						<td class="style13" id="tdDescTable3" style="display: none; "><%=getadminMenuGroupsLngStr("LtxtDescName")%></td>
						<td class="style13"><%=getadminMenuGroupsLngStr("DtxtOrder")%></td>
						<td class="style13">&nbsp;</td>
					</tr>
					<tr>
						<td class="style3">
						<select size="1" name="cmbType" onchange="javascript:changeType(this.value);">
						<option value="S"><%=getadminMenuGroupsLngStr("DtxtSystem")%></option>
						<option value="U"><%=getadminMenuGroupsLngStr("DtxtUDF")%></option>
						<option value="T"><%=getadminMenuGroupsLngStr("DtxtUDT")%></option>
						<option value="F"><%=getadminMenuGroupsLngStr("DtxtField")%></option>
						<option value="Q"><%=getadminMenuGroupsLngStr("LtxtProp")%></option>
						</select></td>
						<td class="style3">
						<input type="button" name="btnQryGroups" id="btnQryGroups" value="..." style="display: none; width: 120px;" onclick="getQryGroups('Add')">
						<input type="hidden" name="QryGroupsAdd" id="QryGroupsAdd" value="">
						<select size="1" name="cmbTableField" onchange="javascript:changeTable(this.value);">
						<option value="OITB"><%=getadminMenuGroupsLngStr("DtxtGroup")%></option>
						<option value="OMRC"><%=getadminMenuGroupsLngStr("DtxtFirm")%></option>
						<option value="OCRD"><%=getadminMenuGroupsLngStr("LtxtSupplier")%></option>
						</select><input type="button" name="btnFormula" disabled id="btnFormulaAddF" value="..." onclick="setFormula('Add', 'F');"></td>
						<td class="style8" id="tdField2" style="display: none; ">
						<select size="1" name="cmbTableFilterID">
						</select></td>
						<td class="style8" id="tdDescTable4" style="display: none; ">
						<select size="1" name="cmbDescTable" id="cmbDescTableAdd" onchange="javascript:changeDescTable(document.frmEditMenuGroup.cmbDescID, document.frmEditMenuGroup.cmbDescName, this.value);" disabled>
						<option></option>
						<%
						rt.movefirst
						do while not rt.eof %>
						<option value="<%=rt(0)%>"><%=rt(0)%></option>
						<% rt.movenext
						loop %>
						</select></td>
						<td class="style8" id="tdDescTable5" style="display: none; ">
						<select size="1" name="cmbDescID" disabled>
						<option value="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</option>
						</select><input type="button" name="btnFormula" id="btnFormulaAddD" disabled value="..." onclick="setFormula('Add', 'D');"></td>
						<td class="style8" id="tdDescTable6" style="display: none; ">
						<select size="1" name="cmbDescName" disabled>
						<option value="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</option>
						</select></td>
						<td class="style8">
							<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td>
									<input type="text" name="AddOrder" id="AddOrder" size="7" style="text-align:right" onfocus="this.select();" onmouseup="event.preventDefault()" onkeydown="return chkMax(event, this, 6);" value="<%=NewOrdr%>">
								</td>
								<td valign="middle">
								<table cellpadding="0" cellspacing="0" border="0">
									<tr>
										<td><img src="images/img_nud_up.gif" id="btnAddOrderUp"></td>
									</tr>
									<tr>
										<td><img src="images/spacer.gif"></td>
									</tr>
									<tr>
										<td><img src="images/img_nud_down.gif" id="btnAddOrderDown"></td>
									</tr>
								</table></td>
							</tr>
						</table>
						<script language="javascript">NumUDAttach('frmEditMenuGroup', 'AddOrder', 'btnAddOrderUp', 'btnAddOrderDown');</script>
						<input type="hidden" name="FilterFormulaAdd" id="AddFilterFormula" value="">
						<input type="hidden" name="DescFormulaAdd" id="AddDescFormula" value="">
						</td>
						<td class="style8">
						<input type="submit" value="<%=getadminMenuGroupsLngStr("DtxtAdd")%>" name="btnAddLine" style="color: #68A6C0; font-family: Verdana; border: 1px solid #68A6C0; background-color: #E5F1FF; font-size:10px; width:75px; height:16px; font-weight:bold" onclick="return valAdd();"></td>
					</tr>
					</table>
				</td>
			</tr>
		<% End If %>
		  </table>
		</td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE" align="center">
		<table border="0" cellpadding="0" width="100%" id="table7">
			<tr>
				<td width="77">
				<input type="submit" value="<%=getadminMenuGroupsLngStr("DtxtApply")%>" name="btnApply" class="OlkBtn" onclick="javascript:return valSave();"></td>
				<td width="77">
				<input type="submit" value="<%=getadminMenuGroupsLngStr("DtxtSave")%>" name="btnSave" class="OlkBtn" onclick="javascript:return valSave();"></td>
				<td><hr size="1"></td>
				<td width="77">
				<input type="button" value="<%=getadminMenuGroupsLngStr("DtxtCancel")%>" name="btnCancel" class="OlkBtn" onclick="javascript:window.location.href='adminMenuGroups.asp'"></td>
			</tr>
		</table>
		</td>
	</tr>
	<input type="hidden" name="cmd" value="<% If Request("editID") <> "" Then %>edit<% ElseIf Request("new") = "Y" Then %>new<% End If %>">
	<input type="hidden" name="editID" value="<%=Request("editID")%>">
	<input type="hidden" name="submitCmd" value="menuGroups">
	<input type="hidden" name="GroupNameTrad" value="">
	<input type="hidden" name="SearchFilterDef" value="">
	</form>
</table>
<script type="text/javascript">
<!--
function Pic(name, page, w, h, s, r) 
{
	var winleft = (screen.width - w) / 2;
	var winUp = (screen.height - h) / 2;
	OpenWin = this.open(page, name, "toolbar=no,menubar=no,location=no,left="+winleft+",top="+winUp+",scrollbars="+s+",resizable="+r+", width="+w+",height="+h);
	//OpenWin.focus()
}

var fldFormula;
function setFormula(FieldID, TypeID)
{
	
	TableID = document.getElementById('cmbDescTable' + FieldID).value;
	
	if (TypeID == 'D' && TableID == '')
	{
		alert('<%=getadminMenuGroupsLngStr("LtxtValDescTable")%>');
		document.getElementById('cmbDescTable' + FieldID).focus();
		return;
	}
	
	Pic('Formula', '', 300, 130, 'N', 'N');
	var Query;
	switch (TypeID)
	{
		case 'F':
			fldFormula = document.getElementById('FilterFormula' + FieldID)
			break;
		case 'D':
			fldFormula  = document.getElementById('DescFormula' + FieldID);
			break;
	}
	doMyLink('adminMenuGroupsFormula.asp', 'pop=Y&Type=' + TypeID + '&TableID=' + TableID + '&Q=' + fldFormula.value, 'Formula');
}
function setFormulaVal(Value)
{
	fldFormula.value = Value;
}
var fldQryGroups;
function getQryGroups(FieldID)
{
	Pic('QryGroups', '', 160, 400, '1', '0');
	
	fldQryGroups = document.getElementById('QryGroups' + FieldID);
	
	doMyLink('selectQryGroups.asp', 'pop=Y&QryGroups=' + fldQryGroups.value, 'QryGroups');
}
function setQryGroups(Value)
{
	fldQryGroups.value = Value;
}

var typeSystem = new Array(new Array('OITB', '<%=getadminMenuGroupsLngStr("DtxtGroup")%>'), new Array('OMRC', '<%=getadminMenuGroupsLngStr("DtxtFirm")%>'), new Array('OCRD', '<%=getadminMenuGroupsLngStr("LtxtSupplier")%>'));

var typeUDF = <%
sql = "select AliasID, Descr " & _
		"from CUFD T0  " & _
		"where TableID = 'OITM' and  " & _
		"( " & _
		"	RTable is not null " & _
		"	or " & _
		"	exists(select '' from UFD1 where TableID = 'OITM' and FieldID = T0.FieldID) " & _
		") "
rs.close
rs.open sql, conn, 3, 1
If Not rs.Eof Then
Response.Write "new Array("
do while not rs.eof
	If rs.bookmark > 1 Then Response.Write ", "
	Response.Write "new Array('" & rs("AliasID") & "', '" & rs("AliasID") & "', '" & Replace(rs("Descr"), "'", "\'") & "')"
rs.movenext
loop
Response.Write ")"
Else
	Response.Write "null"
End If %>;

var typeUDT = <%
sql = "select T0.TableName, T0.[Descr] " & _  
"from OUTB T0 " & _  
"inner join CUFD T1 on T1.TableID = N'@' + T0.TableName " & _  
"where T1.AliasID = 'ItemCode' " 
rs.close
rs.open sql, conn, 3, 1
If Not rs.Eof Then 
	Response.Write "new Array("
	do while not rs.eof
		If rs.bookmark > 1 Then Response.Write ", "
		Response.Write "new Array('" & rs("TableName") & "', '" & Replace(rs("Descr"), "'", "\'") & "')"
	rs.movenext
	loop
	Response.Write ")"
Else
	Response.Write "null"
End If %>;

var typeUDTF = <%
sql = "select Right(X0.TableID, Len(X0.TableID)-1) TableID, X0.AliasID, X0.Descr  " & _  
"from CUFD X0 " & _  
"where X0.TableID in ( " & _  
"	select '@' + T0.TableName " & _  
"	from OUTB T0 " & _  
"	inner join CUFD T1 on T1.TableID = N'@' + T0.TableName " & _  
"	where T1.AliasID = 'ItemCode') " & _  
"and X0.AliasID <> 'ItemCode' " 
rs.close
rs.open sql, conn, 3, 1
If Not rs.Eof Then
	Response.Write "new Array("
	do while not rs.eof
		If rs.bookmark > 1 Then Response.Write ", "
		Response.Write "new Array('" & rs("TableID") & "', '" & rs("AliasID") & "', '" & Replace(rs("Descr"), "'", "\'") & "')"
	rs.movenext
	loop
	Response.Write ")"
Else
	Response.Write "null"
End If %>;

var typeField = new Array(<%
sql = "select name from syscolumns where id = (select id from sysobjects where name = 'OITM')"
rs.close
rs.open sql, conn, 3, 1
do while not rs.eof
	If rs.bookmark > 1 Then Response.Write ","
	Response.Write "'" & rs("name") & "'"
rs.movenext
loop %>);


function changeType(value)
{
	for (var i = document.frmEditMenuGroup.cmbTableField.length-1;i>=0;i--)
	{
		document.frmEditMenuGroup.cmbTableField.remove(i);
	}
	var a;
	var desc;
		
	switch (value)
	{
		case 'S':
			a = typeSystem;
			desc = "<%=getadminMenuGroupsLngStr("LtxtTable")%>";
			document.getElementById('btnFormulaAddF').disabled = true;
			document.getElementById('btnFormulaAddD').disabled = true;
			break;
		case 'U':
			a = typeUDF;
			desc = "<%=getadminMenuGroupsLngStr("DtxtField")%>";
			document.getElementById('btnFormulaAddF').disabled = true;
			document.getElementById('btnFormulaAddD').disabled = true;
			break;
		case 'F':
			a = typeField;
			desc = "<%=getadminMenuGroupsLngStr("DtxtField")%>";
			document.getElementById('btnFormulaAddF').disabled = false;
			document.getElementById('btnFormulaAddD').disabled = false;
			break;
		case 'T':
			a = typeUDT;
			desc = '<%=getadminMenuGroupsLngStr("LtxtTable")%>';
			document.getElementById('btnFormulaAddF').disabled = true;
			document.getElementById('btnFormulaAddD').disabled = true;
			break;
		case 'Q':
			desc = "<%=getadminMenuGroupsLngStr("LtxtProp")%>";
			document.getElementById('btnFormulaAddF').disabled = true;
			document.getElementById('btnFormulaAddD').disabled = true;
			break;
	}
	if (a != null)
	{
		for (var i = 0;i<a.length;i++)
		{
			var opt;
			if (value != 'F')
				opt = new Option(a[i][1],a[i][0]);
			else
				opt = new Option(a[i], a[i]);
			document.frmEditMenuGroup.cmbTableField.options[i] = opt;
		}
	}
	
	txtTblFld.innerHTML = desc
	
	enableDescTable(value == 'F');
	enableTableFilterID(value == 'T');
	
	if (value == 'T')
	{
		changeTable(document.frmEditMenuGroup.cmbTableField.value);
	}
	
	
	document.frmEditMenuGroup.cmbTableField.style.display = (value == 'Q' ? 'none' : '');
	document.frmEditMenuGroup.btnFormulaAddF.style.display = (value == 'Q' ? 'none' : '');
	document.frmEditMenuGroup.btnQryGroups.style.display = (value == 'Q' ? '' : 'none');
}

function enableTableFilterID(enable)
{
	for (var i = 1;i<=2;i++)
	{
		document.getElementById('tdField' + i).style.display = enable ? '' : 'none';
	}
}

function enableDescTable(enable)
{
	document.frmEditMenuGroup.cmbDescTable.disabled = !enable;
	document.frmEditMenuGroup.cmbDescID.disabled = !enable;
	document.frmEditMenuGroup.cmbDescName.disabled = !enable;
	
	for (var i = 1;i<=6;i++)
	{
		document.getElementById('tdDescTable' + i).style.display = enable ? '' : 'none';
	}
}

function changeTable(value)
{
	if (document.frmEditMenuGroup.cmbType.value == 'T')
	{
		for (var i = document.frmEditMenuGroup.cmbTableFilterID.length-1;i>=0;i--)
		{
			document.frmEditMenuGroup.cmbTableFilterID.remove(i);
		}
		
		var curO = 0;
		var found = false;
		for (var i = 0;i<typeUDTF.length;i++)
		{
			if (typeUDTF[i][0] == value)
			{
				found = true
				var opt = new Option(typeUDTF[i][2], typeUDTF[i][1]);
				document.frmEditMenuGroup.cmbTableFilterID.options[curO++] = opt;
			}
			else if (found)
			{
				break;
			}
		}
	}
}

var varDescID;
var varDescName;
function changeDescTable(cmbDescID, cmbDescName, value)
{
	varDescID = cmbDescID;
	varDescName = cmbDescName;
	
	var url='adminMenuGroupsFetch.asp?TableID=' + value;

	xmlHttp=GetXmlHttpObject(setDescTableCols);
	xmlHttp.open("GET", url , true);
	xmlHttp.send(null);
}
function setDescTableCols()
{
	
	if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete")
	{
		var strCols = xmlHttp.responseText.split(', ');
		
		for (var i = varDescID.length-1;i>=1;i--)
		{
			varDescID.remove(i);
			varDescName.remove(i);
		}
		
		if (strCols != '')
		{
			var arrCols = strCols;
			for (var i = 0;i<arrCols.length;i++)
			{
				varDescID.options[i+1] = new Option(arrCols[i], arrCols[i]);
				varDescName.options[i+1] = new Option(arrCols[i], arrCols[i]);
			}
		}
	}
}
function valAdd()
{
	if (document.frmEditMenuGroup.cmbDescTable.selectedIndex > 0 && document.frmEditMenuGroup.cmbType.value == 'F')
	{
		if (document.frmEditMenuGroup.cmbDescID.selectedIndex == 0)
		{
			alert('<%=getadminMenuGroupsLngStr("LtxtValDescID")%>');
			document.frmEditMenuGroup.cmbDescID.focus();
			return false;
		}
		else if (document.frmEditMenuGroup.cmbDescName.selectedIndex == 0)
		{
			alert('<%=getadminMenuGroupsLngStr("LtxtValDescName")%>');
			document.frmEditMenuGroup.cmbDescName.focus();
			return false;
		}
	}
	else if (document.frmEditMenuGroup.cmbType.value == 'Q' && document.frmEditMenuGroup.QryGroupsAdd.value == '')
	{
		alert('<%=getadminMenuGroupsLngStr("LtxtValQryGroups")%>');
		document.frmEditMenuGroup.btnQryGroups.click();
		return false;
	}
	return true;
}
function valSave()
{
	var valDesc = '<%=valDesc%>';
	if (valDesc != '')
	{
		var arrVal = valDesc.split(', ');
		
		for (var i = 0;i<arrVal.length;i++)
		{
			if (document.getElementById('cmbDescTable' + arrVal[i]).selectedIndex > 0)
			{
				if (document.getElementById('cmbDescID' + arrVal[i]).selectedIndex == 0)
				{
					alert('<%=getadminMenuGroupsLngStr("LtxtValDescID")%>');
					document.getElementById('cmbDescID' + arrVal[i]).focus();
					return false;
				}
				else if (document.getElementById('cmbDescName' + arrVal[i]).selectedIndex == 0)
				{
					alert('<%=getadminMenuGroupsLngStr("LtxtValDescName")%>');
					document.getElementById('cmbDescName' + arrVal[i]).focus();
					return false;
				}
			}
		}
	}
	return true;
}
//-->
</script>
<% End If %><!--#include file="bottom.asp" -->
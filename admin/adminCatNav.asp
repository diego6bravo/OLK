<!--#include file="top.asp" -->
<!--#include file="lang/adminCatNav.asp" -->
<!--#include file="adminTradSubmit.asp"-->
<% conn.execute("use [" & Session("olkdb") & "]")

sql = "select NavIndexByX, NavIndexByY from OLKCommon"
set rs = conn.execute(sql)
NavIndexByX = rs("NavIndexByX")
NavIndexByY = rs("NavIndexByY")
rs.close %>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style>

.tdIndex     { font-family: Verdana; font-size: 10px; color: #31659C; }
.cmbSec		 { font-size: 10px; font-family: Verdana; color: #3F7B96; font-weight: bold; border: 1px solid #68A6C0; background-color: #D9F0FD }
.style2 {
	font-weight: normal;
}
.style3 {
	background-color: #F3FBFE;
}
.style4 {
	font-weight: bold;
	background-color: #F3FBFE;
}
.style5 {
	font-family: Verdana;
	font-size: xx-small;
}
.style6 {
	font-family: Verdana;
	font-size: xx-small;
	color: #4783C5;
}
.style7 {
	color: #4783C5;
}
.style8 {
	text-decoration: none;
}
.style9 {
	color: #31659C;
}
.style10 {
	font-family: Verdana;
	font-size: xx-small;
	color: #31659C;
}
.style11 {
	font-family: Verdana;
	font-weight: normal;
	font-size: xx-small;
	color: #31659C;
}
.style12 {
	font-weight: normal;
	color: #31659C;
}
</style>
</head>
<script type="text/javascript" src="scr/calendar.js"></script>
<script type="text/javascript" src="scr/lang/calendar-<%=Left(Session("myLng"), 2)%>.js"></script>
<script type="text/javascript" src="scr/calendar-setup.js"></script>

<table border="0" cellpadding="0" width="100%">
	<% If Request("New") <> "Y" and Request("editIndex") = "" Then %>
	<script language="javascript">
	var txtValNavIndex = '<%=getadminCatNavLngStr("LtxtValNavIndex")%>';
	var txtValEqIndex = '<%=getadminCatNavLngStr("LtxtValEqIndex")%>';
	</script>
	<script language="javascript" src="adminCatNav.js"></script>
	<tr>
		<td bgcolor="#E1F3FD"><span class="style10"><strong>&nbsp;</strong></span><font face="Verdana" color="#31659c" size="1"><strong><%=getadminCatNavLngStr("LttlClientMainNav")%></strong></font></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1" color="#3066E4">
		</font><font face="Verdana" size="1" color="#4783C5"><%=getadminCatNavLngStr("LttlClientMainNavDesc")%></font></p>
		</td>
	</tr>
	<% sql = "select T0.NavIndex, T0.NavTitle, T1.ID OrderID " & _
			"from OLKCatNav T0 " & _
			"left outer join OLKCatNavIndex T1 on T1.NavIndex = T0.NavIndex " & _
			"where T0.Access in ('A', 'C') and T0.Active = 'Y' and T0.NavType = 'M' " & _
			"order by T0.NavTitle asc"
	rs.open sql, conn, 3, 1 %>
	<form method="POST" action="adminSubmit.asp" name="frmNavIndex" onsubmit="return valFrmIndex();">
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" bgcolor="#E2F3FC">
				<tr>
					<td>
					<table border="0" cellspacing="0" width="300">
						<tr>
							<td class="style9"><font face="Verdana" size="1">
							<strong><%=getadminCatNavLngStr("LtxtRows")%></strong></font></td>
							<td><select size="1" name="NavIndexByY" class="input" onchange="doIndex();">
							<% For i = 1 to 10 %>
							<option <% If NavIndexByY = i Then %>selected<% End If %> value="<%=i%>"><%=i%></option>
							<% Next %>
							</select></td>
							<td class="style9"><font face="Verdana" size="1">
							<strong><%=getadminCatNavLngStr("LtxtCols")%></strong></font></td>
							<td><select size="1" name="NavIndexByX" class="input" onchange="doIndex();">
							<% For i = 1 to 4 %>
							<option <% If NavIndexByX = i Then %>selected<% End If %> value="<%=i%>"><%=i%></option>
							<% Next %>
							</select></td>
						</tr>
					</table>
					</td>
				</tr>
				<tr>
					<td id="tdIndex">
					<table cellpadding="0" cellspacing="2" border="0" width="500">
						<% x = 0
						For i = 1 to NavIndexByY %>
						<tr>
							<% For j = 1 to NavIndexByX
								x = x + 1 %>
							<td width="33%" align="center" class="tdIndex">
								<b><%=x%></b><br>
								<select name="NavID<%=x%>" id="cmbNavID" size="1" class="input">
								<option value="-1"><%=getadminCatNavLngStr("LtxtRandom")%></option>
								<% If rs.recordcount > 0 Then rs.movefirst
								do while not rs.eof %>
								<option <% If rs("OrderID") = x Then %>selected<% End If %> value="<%=rs("NavIndex")%>"><%=myHTMLEncode(rs("NavTitle"))%></option>
								<% rs.movenext
								loop %>
								</select>
							</td>
							<% Next %>
						</tr>
						<% Next
						 %>
					</table>
				</td>
				</tr>
				<script language="javascript">
				var myNav = '-1|<%=getadminCatNavLngStr("LtxtRandom")%><% If rs.RecordCount > 0 Then
				rs.movefirst
				do while not rs.eof %>{S}<%=rs("NavIndex")%>|<%=Replace(rs("NavTitle"), "'", "\'")%><% rs.movenext
				loop
				End If
				rs.close %>'.split('{S}');
				</script>
				<input type="hidden" name="cmd" value="NavIndex">
				<input type="hidden" name="submitCmd" value="adminNavCat">
		</table>
		</td>
	</tr>
		<tr>
			<td>
			<table border="0" cellpadding="0" width="100%" id="table22">
				<tr>
					<td width="77">
					<input type="submit" value="<%=getadminCatNavLngStr("DtxtSave")%>" name="btnSaveItemDispORDR" style="color: #68A6C0; font-family: Tahoma; border: 1px solid #68A6C0; background-color: #E5F1FF; font-size:10px; width:75; height:22; font-weight:bold"></td>
					<td><hr color="#0D85C6" size="1"></td>
				</tr>
			</table>
			</td>
		</tr>
		</form>
		<tr>
			<td bgcolor="#E1F3FD"><b><font color="#FFFFFF">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><strong><%=getadminCatNavLngStr("LttlNavLst")%></strong></font></b></td>
		</tr>
		<tr>
			<td bgcolor="#F5FBFE">
			<p align="justify"><img src="images/lentes.gif">
			<font face="Verdana" size="1" color="#4783C5"><%=getadminCatNavLngStr("LttlNavDesc")%></font></p>
			</td>
		</tr>
		<form method="post" action="adminSubmit.asp" name="frmPolls">
		<%
		sql = "select NavIndex, NavTitle, NavDesc, NavType, Case NavType When 'M' Then 1 When 'S' Then 2 When 'Q' Then 3 End NavTypeOrdr, " & _
		"Active, Access from OLKCatNav " & _
		"order by NavTypeOrdr, NavTitle"
		'rs.close
	    rs.open sql, conn, 3, 1
		If rs.recordcount > 0 then %>
		<tr>
			<td>
			<table border="0" cellpadding="0" width="100%" id="table12">
				<tr>
					<td align="center" bgcolor="#E2F3FC" width="10">&nbsp;</td>
					<td align="center" bgcolor="#E2F3FC" class="style10">
					<font size="1" face="Verdana" color="#31659C"><strong><%=getadminCatNavLngStr("DtxtType")%>
					</strong>
					</font></td>
					<td align="center" bgcolor="#E2F3FC">
					<font size="1" face="Verdana" color="#31659C"><strong><%=getadminCatNavLngStr("DtxtTitle")%>
					</strong>
					</font></td>
					<td align="center" bgcolor="#E2F3FC">
					<font size="1" face="Verdana" color="#31659C"><strong><%=getadminCatNavLngStr("DtxtDescription")%>
					</strong>
					</font></td>
					<td align="center" bgcolor="#E2F3FC">
					<font size="1" face="Verdana" color="#31659C"><strong><%=getadminCatNavLngStr("DtxtAccess")%>
					</strong>
					</font></td>
					<td align="center" bgcolor="#E2F3FC">
					<font size="1" face="Verdana" color="#31659C"><strong><%=getadminCatNavLngStr("DtxtActive")%>
					</strong>
					</font></td>
					<td align="center" bgcolor="#E2F3FC" width="16">&nbsp;</td>
				</tr>
				<script language="javascript">
				function delNav(navIndex, navTitle)
				{
					if(confirm('<%=getadminCatNavLngStr("LtxtConfDelNav")%>'.replace('{0}', navTitle)))
						window.location.href = 'adminSubmit.asp?cmd=del&delIndex=' + navIndex + '&submitCmd=adminNavCat';
				}
				</script>
				<%
				do While NOT RS.EOF 
			   varx = varx + 1 %>
			   <input type="hidden" name="NavIndex" value="<%=rs("NavIndex")%>">
				<tr bgcolor="#F3FBFE">
					<td width="10" valign="top">
					<a href="adminCatNav.asp?editIndex=<%=rs("NavIndex")%>">
					<img border="0" src="images/<%=Session("rtl")%>flechaselec.gif"></a></td>
					<td valign="top" class="style6">
					<span class="style7">
					<font size="1" face="Verdana"><%
					Select Case RS("NavType")
						Case "M" %><%=getadminCatNavLngStr("LtxtMain")%>
					<%	Case "S" %><%=getadminCatNavLngStr("LtxtSubNav")%>
					<%	Case "Q" %><%=getadminCatNavLngStr("DtxtQuery")%>
					<% End Select %></font></span><span class="style6"> </span>
					</td>
					<td valign="top">
					<table cellpadding="0" cellspacing="0" border="0" width="100%">
						<tr>
							<td class="style7"><font size="1" face="Verdana"><%=RS("NavTitle")%></font></td>
							<td width="16" class="style5">
							<a href="javascript:doFldTrad('CatNav', 'NavIndex', <%=rs("NavIndex")%>, 'AlterNavTitle', 'T', null);" class="style8">
							<span class="style7"><img src="images/trad.gif" alt="<%=getadminCatNavLngStr("DtxtTranslate")%>" border="0"></span></a></td>
						</tr>
					</table>
					</td>
					<td valign="top">
					<table cellpadding="0" cellspacing="0" border="0" width="100%">
						<tr>
							<td class="style7"><font size="1" face="Verdana"><%=Left(RS("NavDesc"), 100)%>&nbsp;</font></td>
							<td width="16" valign="top" class="style5">
							<a href="javascript:doFldTrad('CatNav', 'NavIndex', <%=rs("NavIndex")%>, 'AlterNavDesc', 'M', null);" class="style8">
							<span class="style7"><img src="images/trad.gif" alt="<%=getadminCatNavLngStr("DtxtTranslate")%>" border="0"></span></a></td>
						</tr>
					</table></td>
					<td valign="top">
					<p align="center" class="style7">
					<font size="1" face="Verdana">
					<%
					Select Case RS("Access")
						Case "C" %><%=getadminCatNavLngStr("DtxtClients")%>
					<%	Case "S" %><%=getadminCatNavLngStr("DtxtAgents")%>
					<%	Case "A" %><%=getadminCatNavLngStr("DtxtAll")%>
					<% End Select %>
					</font></td>
					<td valign="top">
					<p align="center" class="style7">
					<input type="checkbox" class="noborder" name="Active<%=rs("NavIndex")%>" value="Y" <% If rs("Active") = "Y" Then %>checked<% End If %>></td>
					<td valign="middle" width="16" valign="top">
					<a href="javascript:delNav(<%=rs("NavIndex")%>, '<%=Replace(Rs("NavTitle"),"'","\'")%>');">
					<img border="0" src="images/remove.gif" width="16" height="16"></a></td>
				</tr>
				<input type="hidden" name="NavIndex" value="<%=rs("NavIndex")%>">
				<% RS.MoveNext
				loop %>
			</table>
			</td>
		</tr>
		<% End If %>
		<tr>
			<td>
			<table border="0" cellpadding="0" width="100%" id="table22">
				<tr>
					<td width="77">
					<input type="submit" value="<%=getadminCatNavLngStr("DtxtSave")%>" name="btnSave" class="OlkBtn"></td>
					<td width="77">
					<input type="button" value="<%=getadminCatNavLngStr("DtxtNew")%>" name="btnNew" class="OlkBtn" onclick="javascript:window.location.href='adminCatNav.asp?New=Y'"></td>
					<td><hr color="#0D85C6" size="1"></td>
				</tr>
			</table>
			</td>
		</tr>
		<input type="hidden" name="submitCmd" value="admCatNav">
		</form>
	<% ElseIf Request("editIndex") <> "" or Request("New") = "Y" Then %>
	<script language="javascript">
	function valFrm2()
	{
		if (document.form2.NavTitle.value == '')
		{
			alert('<%=getadminCatNavLngStr("LtxtValTitle")%>');
			document.form2.NavTitle.focus();
			return false;
		}
		else if (document.form2.NavImgTypeQ.checked && document.form2.NavImgQry.value == '')
		{
			alert('<%=getadminCatNavLngStr("LtxtValImgQry")%>');
			document.form2.NavImgQry.focus();
			return false;
		}
		else if (document.form2.NavImgTypeQ.checked && document.form2.valImgQry.value == 'Y')
		{
			alert('<%=getadminCatNavLngStr("LtxtValImgQryVal")%>');
			document.form2.btnVerfyImgQry.focus();
			return false;
		}
		else if (document.form2.NavType.value == 'Q' && document.form2.NavQry.value == '')
		{
			alert('<%=getadminCatNavLngStr("LtxtValItmQry")%>');
			document.form2.NavQry.focus();
			return false;
		}
		else if (document.form2.NavType.value == 'Q' && document.form2.valNavQry.value == 'Y')
		{
			alert('<%=getadminCatNavLngStr("LtxtValItmQryVal")%>')
			document.form2.btnVerfy.focus();
			return false;
		}
		setSubIndex();
		return true;
	}
	function setSubIndex()
	{
		var retVal = '';
		for (var i = 0;i<document.form2.NavListAdd.length;i++)
		{
			if (retVal != '') retVal += ', ';
			retVal += document.form2.NavListAdd.options[i].value;
		}
		document.form2.SubIndex.value = retVal;
	}
	</script>
	<form method="POST" action="adminsubmit.asp" name="form2" onsubmit="javascript:return valFrm2()">
		<tr>
			<td bgcolor="#E1F3FD">&nbsp;<b><font size="1" face="Verdana" color="#31659C"><% If Request("editIndex") = "" Then %><%=getadminCatNavLngStr("DtxtAdd")%><% Else %><%=getadminCatNavLngStr("LtxtEdit")%><% End If %>&nbsp;<%=getadminCatNavLngStr("LtxtNav")%></font></b></td>
		</tr>
		<tr>
			<td bgcolor="#F5FBFE">
			<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1">
			<font color="#4783C5"><%=getadminCatNavLngStr("LtxtAddEditDesc")%></font></font></p>
			</td>
		</tr>
		<tr>
			<td>
			<% 
			If Request("editIndex") <> "" Then
				sql = 	"select NavTitle, IsNull(NavDesc, '') NavDesc, NavImg, NavQry, ShowFrom, ShowTo, " & _
						"Active, Access, NavType, CatType, NavImgType, NavImgQry, AutoRedir, ApplyAnonCatFilter " & _
						"from OLKCatNav " & _
						"where NavIndex = " & Request("editIndex") 
				set rs = conn.execute(sql) 
				NavTitle = rs("NavTitle")
				NavDesc = rs("NavDesc")
				NavImg = rs("NavImg")
				NavQry = rs("NavQry")
				If Not IsNull(NavImg) Then showImg = NavImg Else showImg = "n_a.gif"
				ShowFrom = rs("ShowFrom")
				ShowTo = rs("ShowTo")
				Active = rs("Active")
				Access = rs("Access")
				NavType = rs("NavType")
				CatType = rs("CatType")
				NavImgType = rs("NavImgType")
				NavImgQry = rs("NavImgQry")
				AutoRedir = rs("AutoRedir")
				ApplyAnonCatFilter = rs("ApplyAnonCatFilter")
				rs.close
			Else
				showImg = "n_a.gif"
				CatType = "C"
				NavType = "M"
				NavImgType = "I"
				AutoRedir = "N"
				NavTitle = "" %>
			<input type="hidden" name="NavTitleTrad">
			<input type="hidden" name="NavDescTrad">
			<input type="hidden" name="NavQryDef">
			<input type="hidden" name="NavImgQryDef">
			<% End If %>
			<table border="0" cellpadding="0" width="100%" id="table20">
				<tr>
					<td>
					<table border="0" cellpadding="0" id="table20">
						<tr>
							<td bgcolor="#E2F3FC" class="style10" style="width: 100px">
							<p class="style9">
							<font face="Verdana" size="1">
							<strong><%=getadminCatNavLngStr("DtxtTitle")%></strong></font></td>
							<td class="style3">
							<table cellpadding="0" cellspacing="0" border="0" width="100%">
								<tr>
									<td valign="bottom"><input style="width:100%;" name="NavTitle" class="input" value="<%=Server.HTMLEncode(NavTitle)%>" size="72" onkeydown="return chkMax(event, this, 50);">
									</td>
									<td width="16"><a href="javascript:doFldTrad('CatNav', 'NavIndex', '<%=Request("editIndex")%>', 'AlterNavTitle', 'T', <% If Request("editIndex") <> "" Then %>null<% Else %>document.form2.NavTitleTrad<% End If %>);"><img src="images/trad.gif" alt="<%=getadminCatNavLngStr("DtxtTranslate")%>" border="0"></a></td>
									<td style="width: 100px" align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><font face="Verdana" size="1" color="#4783C5"><b>
									<input type="checkbox" <% If Active = "Y" Then %>checked<% End If %> name="Active" class="noborder" value="Y" id="Active"></b><label for="Active"><%=getadminCatNavLngStr("DtxtActive")%></label></font></td>
								</tr>
							</table>
							</td>
						</tr>
						<tr>
							<td bgcolor="#E2F3FC" valign="top" class="style10" style="width: 100px">
							<p><span class="style12">
							<font face="Verdana" size="1"><strong><%=getadminCatNavLngStr("DtxtDescription")%></strong></font></span><span class="style11"><strong>
							</strong></span>
							</p></td>
							<td width="540" valign="bottom" class="style3">
							<table cellpadding="0" cellspacing="0" border="0">
								<tr>
									<td><textarea rows="4" name="NavDesc" cols="83" class="input"><%=Server.HTMLEncode(NavDesc)%></textarea></td>
									<td valign="bottom"><a href="javascript:doFldTrad('CatNav', 'NavIndex', '<%=Request("editIndex")%>', 'AlterNavDesc', 'M', <% If Request("editIndex") <> "" Then %>null<% Else %>document.form2.NavDescTrad<% End If %>);"><img src="images/trad.gif" alt="<%=getadminCatNavLngStr("DtxtTranslate")%>" border="0"></a></td>
								</tr>
							</table>
							</td>
						</tr>
						<tr>
							<td bgcolor="#E2F3FC" class="style10" style="width: 100px">
							<p><span class="style12">
							<font face="Verdana" size="1"><strong><%=getadminCatNavLngStr("DtxtAccess")%></strong></font></span><span class="style11"><strong>
							</strong></span>
							</p>
							</td>
							<td valign="top" width="540" class="style3">
							<select size="1" name="Access">
							<option value="A"><%=getadminCatNavLngStr("DtxtAll")%></option>
							<option <% If Access = "C" Then %>selected<% End If %> value="C"><%=getadminCatNavLngStr("DtxtClients")%></option>
							<option <% If Access = "V" Then %>selected<% End If %> value="V"><%=getadminCatNavLngStr("DtxtAgents")%></option>
							</select> </td>
						</tr>
						<tr>
							<td valign="top" bgcolor="#E2F3FC" class="style10" style="width: 100px">
							<p class="style9">
							<font face="Verdana" size="1"><strong><%=getadminCatNavLngStr("LtxtImageType")%>
							</strong>
							</font>
							</td>
							<td valign="top" width="540" class="style3">
							<font face="Verdana" size="1" color="#4783C5">
							<b>
							<input type="radio" class="noborder" <% If NavImgType = "I" Then %>checked<% End If %> name="NavImgType" id="NavImgTypeI" value="I" onclick="changeImgType(this.value);"></b><label for="NavImgTypeI"><%=getadminCatNavLngStr("DtxtImage")%></label></font> <font face="Verdana" size="1" color="#4783C5">
							<b>
							<input type="radio" class="noborder" <% If NavImgType = "Q" Then %>checked<% End If %> name="NavImgType" id="NavImgTypeQ" value="Q" onclick="changeImgType(this.value);"></b><label for="NavImgTypeQ"><%=getadminCatNavLngStr("DtxtQuery")%></label></font></td>
						</tr>
						<tr id="trImgImg" <% If NavImgType = "Q" Then %>style="display: none;"<% End If %>>
							<td valign="top" bgcolor="#E2F3FC" height="80" class="style10" style="width: 100px">
							<p class="style9">
							<font face="Verdana" size="1"><strong><%=getadminCatNavLngStr("DtxtImage")%>
							</strong>
							</font>
							</td>
							<td valign="bottom" width="540" height="80" class="style3">
							<table border="0" id="table7" cellpadding="0">
							<tr>
								<td>
								<img border="1" name="navImage" id="navImage" src="pic.aspx?FileName=<%=showImg%>&MaxSize=80&dbName=<%=Session("olkdb")%>"></td>
								<td valign="bottom">
								<input type="button" value="<%=getadminCatNavLngStr("LtxtUpload")%>" name="B2" class="OlkBtn" onclick="javascript:Start('upload/fileupload.aspx?ID=<%=Session("ID")%>&style=admin/style/style_pop.css&Source=Admin',300,100,'no')">
								<input type="button" value="X" name="btnRemImg" <% If showImg = "n_a.gif" Then %>disabled<% End If %> style="color: #68A6C0; font-family: Tahoma; border: 1px solid #68A6C0; background-color: #E5F1FF; font-size:10px; width:27; height:23; font-weight:bold" onclick="javascript:if(confirm('<%=getadminCatNavLngStr("LtxtConfRemImg")%>')){document.form2.NavImg.value='';document.form2.navImage.src='pic.aspx?FileName=n_a.gif&MaxSize=80&dbName=<%=Session("dbName")%>';this.disabled=true;}">
								</td>
							</tr>
							</table>
							</td>
						</tr>
						<tr id="trImgQry" <% If NavImgType = "I" Then %>style="display: none;"<% End If %>>
							<td bgcolor="#E2F3FC" height="80" valign="top" class="style10" style="width: 100px">
							<p class="style9">
							<font face="Verdana" size="1"><strong><%=getadminCatNavLngStr("LtxtImgQry")%></strong></font></td>
							<td width="540" height="80" class="style3">
							<table cellpadding="0" cellspacing="0" border="0" width="100%">
								<tr>
									<td rowspan="2">
										<textarea dir="ltr" rows="6" name="NavImgQry" id="NavImgQry" cols="82" style="width: 100%;" class="input" onkeypress="javascript:document.form2.btnVerfyImgQry.src='images/btnValidate.gif';document.form2.btnVerfyImgQry.style.cursor = 'hand';document.form2.valImgQry.value='Y';"><%=myHTMLEncode(NavImgQry)%></textarea>
									</td>
									<td valign="top" width="1">
										<img style="cursor: hand; " src="images/qry_note.gif" id="btnNoteFilter" alt="<%=getadminCatNavLngStr("DtxtDefinition")%>" onclick="javascript:doFldNote(17, 'NavImgQry', '<%=Request("editIndex")%>', <% If Request("editIndex") <> "" Then %>null<% Else %>document.form2.NavImgQryDef<% End If %>);">
									</td>
								</tr>
								<tr>
									<td valign="bottom" width="1">
										<img src="images/btnValidateDis.gif" id="btnVerfyImgQry" alt="|D:txtValidate|" onclick="javascript:if (document.form2.valImgQry.value == 'Y')VerfyQuery(this, document.form2.valImgQry, 'NavImgQry');">
										<input type="hidden" name="valImgQry" id="valImgQry" value="N">
									</td>
								</tr>
							</table>
							</td>
						</tr>
						<tr id="trImgQryVars" <% If NavImgType = "I" Then %>style="display: none;"<% End If %>>
							<td valign="top" bgcolor="#E2F3FC" class="style10" style="width: 100px">
							<p class="style9">
									<font size="1" face="Verdana">
									<strong><%=getadminCatNavLngStr("DtxtVariables")%></strong></font></td>
							<td valign="top" width="540" class="style3">
									<font size="1" color="#4783C5" face="Verdana">
									<span dir="ltr">@CardCode</span> = <%=getadminCatNavLngStr("DtxtClientCode")%><br>
									<span dir="ltr">@SlpCode</span> = <%=getadminCatNavLngStr("DtxtAgentCode")%></font></td>
						</tr>
						<tr>
							<td bgcolor="#E2F3FC" class="style10" style="width: 100px">
							<p class="style9">
							<font face="Verdana" size="1"><strong><%=getadminCatNavLngStr("DtxtDate")%></strong></font></td>
							<td width="540" class="style3">
							<table cellpadding="0" cellspacing="0" border="0">
								<tr>
									<td>
									<p align="right"><font face="Verdana" size="1" color="#4783C5"><%=getadminCatNavLngStr("DtxtFrom")%></font></td>
									<td width="16">
									<img border="0" src="images/cal.gif" id="btnShowFrom" width="16" height="16" style="float:left;padding-left:1px;padding-top:1px"></td>
									<td><input readonly class="input" type="text" name="ShowFrom" id="ShowFrom" size="11" value="<%=FormatDate(ShowFrom, False)%>" onclick="btnShowFrom.click()"></td>
									<td><img border="0" src="images/remove.gif" width="16" height="16" onclick="ShowFrom.value='';"></td>
									<td width="10"><font size="1">&nbsp;</font></td>
									<td>
									<p align="right">
									<font face="Verdana" size="1" color="#4783C5"><%=getadminCatNavLngStr("DtxtTo")%></font></td>
									<td width="16">
									<img border="0" src="images/cal.gif" id="btnShowTo" width="16" height="16" style="float:left;padding-left:1px;padding-top:1px"></td>
									<td><input readonly class="input"  type="text" name="ShowTo" id="ShowTo" size="11" value="<%=FormatDate(ShowTo, False)%>" onclick="btnShowTo.click()"></td>
									<td><img border="0" src="images/remove.gif" width="16" height="16" onclick="ShowTo.value='';"></td>
								</tr>
							</table>
							</td>
						</tr>
						<tr>
							<td valign="top" bgcolor="#E2F3FC" class="style10" style="width: 100px">
							<p class="style9">
							<font face="Verdana" size="1"><strong><%=getadminCatNavLngStr("DtxtType")%></strong></font></td>
							<td valign="top" width="540" class="style3">
									<select size="1" name="NavType" onchange="changeNav(this.value);">
									<option value="M"><%=getadminCatNavLngStr("LtxtMain")%></option>
									<option <% If NavType = "S" Then %>selected<% End If %> value="S"><%=getadminCatNavLngStr("LtxtSubNav")%></option>
									<option <% If NavType = "Q" Then %>selected<% End If %> value="Q"><%=getadminCatNavLngStr("DtxtQuery")%></option>
									</select> 
									<font face="Verdana" size="1" color="#4783C5">
									<b>
									<input type="checkbox" <% If NavType <> "M" Then %>disabled<% End If %> <% If AutoRedir = "Y" Then %>checked<% End If %> class="noborder" name="AutoRedir" value="Y" id="AutoRedir"></b><label for="AutoRedir"><%=getadminCatNavLngStr("LtxtReditAuto")%></label></font></td>
						</tr>
						<tr id="trQry1" <% If NavType <> "Q" Then %>style="display: none"<% End If %>>
							<td valign="top" bgcolor="#E2F3FC" class="style10" style="width: 100px">
							<p class="style9">
							<font face="Verdana" size="1"><strong><%=getadminCatNavLngStr("LtxtItemsView")%></strong></font></td>
							<td valign="top" width="540" class="style4">
									<font face="Verdana" size="1" color="#4783C5">
									<input class="noborder" type="radio" name="CatType" <% If CatType = "C" Then %>checked<% End If %> value="C" id="CatTypeC"><span class="style2"><label for="CatTypeC"><%=getadminCatNavLngStr("DtxtCat")%></label></span>
									<input class="noborder" type="radio" value="S" <% If CatType = "S" Then %>checked<% End If %> name="CatType" id="CatTypeS"><span class="style2"><label for="CatTypeS"><%=getadminCatNavLngStr("DtxtStore")%></label></span>
									<input class="noborder" type="radio" value="L" <% If CatType = "L" Then %>checked<% End If %> name="CatType" id="CatTypeL"><span class="style2"><label for="CatTypeL"><%=getadminCatNavLngStr("DtxtList")%></label></span></font></td>
						</tr>
						<tr id="trQry2" <% If NavType <> "Q" Then %>style="display: none"<% End If %>>
							<td valign="top" bgcolor="#E2F3FC" class="style10" style="width: 100px">
							<p>
							<font face="Verdana" size="1" class="style9">
							<strong><%=getadminCatNavLngStr("DtxtQuery")%><br>
							(</strong>ItemCode in<strong>)</strong></font></td>
							<td valign="top" width="540" class="style3">
							<table cellpadding="0" cellspacing="0" border="0" width="100%">
								<tr>
									<td rowspan="2">
										<textarea dir="ltr" rows="10" name="NavQry" id="NavQry" cols="82" style="width: 100%;" class="input" onkeypress="javascript:document.form2.btnVerfy.src='images/btnValidate.gif';document.form2.btnVerfy.style.cursor = 'hand';document.form2.valNavQry.value='Y';"><%=myHTMLEncode(NavQry)%></textarea>
									</td>
									<td valign="top" width="1">
										<img style="cursor: hand; " src="images/qry_note.gif" id="btnNoteFilter" alt="<%=getadminCatNavLngStr("DtxtDefinition")%>" onclick="javascript:doFldNote(17, 'NavQry', '<%=Request("editIndex")%>', <% If Request("editIndex") <> "" Then %>null<% Else %>document.form2.NavQryDef<% End If %>);">
									</td>
								</tr>
								<tr>
									<td valign="bottom" width="1">
										<img src="images/btnValidateDis.gif" id="btnVerfy" alt="|D:txtValidate|" onclick="javascript:if (document.form2.valNavQry.value == 'Y')VerfyQuery(this, document.form2.valNavQry, 'NavQry');">
										<input type="hidden" name="valNavQry" id="valNavQry" value="N">
									</td>
								</tr>
							</table>
							</td>
						</tr>
						<tr id="trQry3" <% If NavType <> "Q" Then %>style="display: none"<% End If %>>
							<td valign="top" bgcolor="#E2F3FC" class="style10" style="width: 100px">
							<p>
									<strong><span class="style9">
									<font size="1" face="Verdana">
									<%=getadminCatNavLngStr("DtxtVariables")%></font></span></strong></td>
							<td valign="top" width="540" class="style3">
									<font size="1" color="#4783C5" face="Verdana">
									<span dir="ltr">@CardCode</span> = <%=getadminCatNavLngStr("DtxtClientCode")%><br>
									<span dir="ltr">@SlpCode</span> = <%=getadminCatNavLngStr("DtxtAgentCode")%></font></td>
						</tr>
						<tr id="trQry4" <% If NavType <> "Q" Then %>style="display: none"<% End If %>>
							<td valign="top" bgcolor="#E2F3FC" class="style10" style="width: 100px">&nbsp;</td>
							<td valign="top" width="540" class="style3">
							<input class="noborder" type="checkbox" name="ApplyAnonCatFilter" <% If ApplyAnonCatFilter = "Y" Then %>checked<% End If %> value="Y" id="ApplyAnonCatFilter">
							<font size="1" color="#4783C5" face="Verdana"><label for="ApplyAnonCatFilter"><%=getadminCatNavLngStr("LtxtApplyAnonCatFilte")%></label></font></td>
						</tr>
						<tr id="trSub" <% If NavType = "Q" Then %>style="display: none"<% End If %>>
							<td bgcolor="#E2F3FC" height="174" valign="top" class="style10" style="width: 100px">
							<p class="style9">
							<font face="Verdana" size="1"><strong><%=getadminCatNavLngStr("LtxtSubNavs")%></strong></font></td>
							<td width="540" valign="top" class="style3">
							<table border="0" cellpadding="0" cellspacing="1" width="100%" id="table23">
								<tr>
									<td>
									<font face="Verdana" size="1" color="#4783C5">
									<%=getadminCatNavLngStr("LtxtSelNavs")%></font></td>
									<td width="100" align="center">
									&nbsp;</td>
								</tr>
							<% sql = 	"select NavIndex, " & _
										"Case NavType When 'M' Then N'" & getadminCatNavLngStr("LtxtMain") & "' When 'S' Then N'" &  getadminCatNavLngStr("LtxtSubNav") & "' When 'Q' Then N'" & getadminCatNavLngStr("DtxtQuery") & "' End + N' - ' + NavTitle NavTitle, " & _
										"Case NavType When 'M' Then 1 When 'S' Then 2 When 'Q' Then 3 End NavTypeOrdr, " & _
										""
								If Request("editIndex") <> "" Then 
									sql = sql & " Case When Exists(select 'A' from OLKCatNavSub where SubIndex = T0.NavIndex and NavIndex = " & Request("editIndex") & ") Then 'Y' Else 'N' End "
								Else
									sql = sql & " 'N' "
								End If
								sql = sql & " Verfy from OLKCatNav T0 "
								If Request("editIndex") <> "" Then sql = sql & " where NavIndex <> " & Request("editIndex") & " "
								sql = sql & "order by NavTypeOrdr "
								rs.open sql, conn, 3, 1 %>
								<tr>
									<td>
									<select size="12" name="NavListAdd" style="width: 440px; height: 107px; ">
									<% rs.Filter = "Verfy = 'Y'"
									do while not rs.eof %>
									<option value="<%=rs("NavIndex")%>"><%=myHTMLEncode(rs("NavTitle"))%></option>
									<% rs.movenext
									loop %>
									</select></td>
									<td width="100" align="center" valign="bottom">
									<input type="button" value="<%=getadminCatNavLngStr("LtxtRemove")%>" name="btnRemSub" class="OlkBtn" onclick="remSub();"></td>
								</tr>
								<tr>
									<td>
									<font face="Verdana" size="1" color="#4783C5">
									<%=getadminCatNavLngStr("LtxtAvlNavs")%></font></td>
									<td width="100" align="center">
									&nbsp;</td>
								</tr>
								<tr>
									<td>
									<select size="1" name="NavList" style="width: 440px; height:23px; ">
									<% rs.Filter = "Verfy = 'N'"
									do while not rs.eof %>
									<option value="<%=rs("NavIndex")%>"><%=myHTMLEncode(rs("NavTitle"))%></option>
									<% rs.movenext
									loop %>
									</select></td>
									<td width="100" align="center">
									<input type="button" value="<%=getadminCatNavLngStr("DtxtAdd")%>" name="btnAddSub" class="OlkBtn" onclick="addSub();"></td>
								</tr>
							</table>
							</td>
						</tr>
						</table>
					</td>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td>
			<table border="0" cellpadding="0" width="100%" id="table5">
				<tr>
					<td width="77">
					<input type="submit" value="<%=getadminCatNavLngStr("DtxtApply")%>" name="btnApply" class="OlkBtn"></td>
					<td width="77">
					<input type="submit" value="<%=getadminCatNavLngStr("DtxtSave")%>" name="btnSave" class="OlkBtn"></td>
					<td><hr color="#0D85C6" size="1"></td>
					<td width="77">
					<input type="button" value="<%=getadminCatNavLngStr("DtxtCancel")%>" name="B2" class="OlkBtn" onclick="javascript:if(confirm('<%=getadminCatNavLngStr("DtxtConfCancel")%>'))window.location.href='adminCatNav.asp'"></td>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td>&nbsp;</td>
		</tr>
		<input type="hidden" name="editIndex" value="<%=Request("editIndex")%>">
		<input type="hidden" name="submitCmd" value="adminNavCat">
		<input type="hidden" name="cmd" value="<% If Request("editIndex") <> "" Then %>edit<% Else %>add<% End If %>">
		<% End If %>
		<input type="hidden" name="NavImg" value="<%=NavImg%>">
		<input type="hidden" name="SubIndex" value="">
	</form>
</table>
<% If Request("editIndex") <> "" or Request("New") = "Y" Then %>
<script language="javascript">
Calendar.setup({
    inputField     :    "ShowFrom",     // id of the input field
    ifFormat       :    "<%=GetCalendarFormatString%>",      // format of the input field
    button         :    "btnShowFrom",  // trigger for the calendar (button ID)
    align          :    "Bl",           // alignment (defaults to "Bl")
    singleClick    :    true
});
Calendar.setup({
    inputField     :    "ShowTo",     // id of the input field
    ifFormat       :    "<%=GetCalendarFormatString%>",      // format of the input field
    button         :    "btnShowTo",  // trigger for the calendar (button ID)
    align          :    "Bl",           // alignment (defaults to "Bl")
    singleClick    :    true
});
function Start(page, w, h, s) {
	OpenWin = this.open(page, "DatePicker", "toolbar=no,menubar=no,location=no,scrollbars="+s+",resizable=no, width="+w+",height="+h);
}
function setTimeStamp(vardate) {
	document.form1.newsDate.value = vardate
}
function changepic(img_src) {
	document['navImage'].src="pic.aspx?FileName="+img_src+'&maxSize=80&dbName=<%=Session("olkdb")%>';
	document.form2.NavImg.value = img_src
	document.form2.btnRemImg.disabled = false;
}

var myBtnVerfy;
var myHdVerfy;
function VerfyQuery(btn, hd, cmd)
{
	myBtnVerfy = btn;
	myHdVerfy = hd;
	document.frmVerfyQuery.Query.value = document.getElementById(cmd).value;
	document.frmVerfyQuery.type.value = cmd;
	document.frmVerfyQuery.submit();
}

function VerfyQueryVerified()
{
	myBtnVerfy.src='images/btnValidateDis.gif'
	myBtnVerfy.cursor = '';
	myHdVerfy.value='N';
}
function changeNav(val)
{
	trQry1.style.display = val == 'Q' ? '' : 'none';
	trQry2.style.display = val == 'Q' ? '' : 'none';
	trQry3.style.display = val == 'Q' ? '' : 'none';
	trQry4.style.display = val == 'Q' ? '' : 'none';
	trSub.style.display = val != 'Q' ? '' : 'none';
	document.form2.AutoRedir.disabled = val == 'M' ? false : true;
}
function changeImgType(val)
{
	trImgImg.style.display = val == 'I' ? '' : 'none';
	trImgQry.style.display = val == 'Q' ? '' : 'none';
	trImgQryVars.style.display = trImgQry.style.display;
}
function addSub()
{
	var sourceSub = document.form2.NavList;
	var targetSub = document.form2.NavListAdd;
	transferSub(sourceSub, targetSub);
}
function remSub()
{
	var sourceSub = document.form2.NavListAdd;
	var targetSub = document.form2.NavList;
	transferSub(sourceSub, targetSub);
}
function transferSub(sourceSub, targetSub)
{
	if (sourceSub.value != '')
	{
		sourceID = sourceSub.value;
		sourceText = sourceSub.options[sourceSub.selectedIndex].text;
		targetSub.options[targetSub.length++] = new Option(sourceText, sourceID);
		if (browserDetect() == 'msie')
		{
			sourceSub.options.remove(sourceSub.selectedIndex);
		}
		else
		{
			sourceSub.remove(sourceSub.selectedIndex);
		}
	}
}
</script>
<form name="frmVerfyQuery" action="verfyQuery.asp" method="post" target="iVerfyQuery">
	<iframe id="iVerfyQuery" name="iVerfyQuery" style="display: none" src>
	</iframe><input type="hidden" name="type" value="NavQry">
	<input type="hidden" name="Query" value>
	<input type="hidden" name="parent" value="Y">
</form>
<% End If %><!--#include file="bottom.asp" -->
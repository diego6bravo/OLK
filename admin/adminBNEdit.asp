<!--#include file="top.asp" -->
<!--#include file="lang/adminBNEdit.asp" -->
<!--#include file="adminTradSubmit.asp"-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style type="text/css">
.style1 {
	background-color: #E2F3FC;
	font-family: Verdana;
	font-size: xx-small;
	color: #31659C;
}
.style2 {
	background-color: #F7FBFF;
}
.style3 {
	font-weight: normal;
}
.style4 {
	color: #31659C;
}
</style>
</head>

<script type="text/javascript" src="scr/calendar.js"></script>
<script type="text/javascript" src="scr/lang/calendar-<%=Left(Session("myLng"), 2)%>.js"></script>
<script type="text/javascript" src="scr/calendar-setup.js"></script>
<%
	conn.execute("use [" & Session("olkdb") & "]")
	
	If Request("BannerID") <> "" Then
		sql = "select BannerID,BannerDesc,Link,Picture,Query ,GroupID, StartDate, EndDate, Status "& _
				 "from OLKBN "& _
				 "WHERE BannerID = "& Request("BannerID")
		Set rs = conn.Execute(sql)
		BannerDesc = rs("BannerDesc")
		Link = rs("Link")
		Picture = rs("Picture")
		Query = rs("Query")
		GroupID = rs("GroupID")
		StartDate = rs("StartDate")
		EndDate = rs("EndDate")
		Status = rs("Status")
		
		hdnCodClientes = ""
		sql = "select GroupCode from OLKBNOCRG where BannerID = " & Request("BannerID")
		set rs = conn.execute(sql)
		do while not rs.eof
			If hdnCodClientes <> "" Then hdnCodClientes = hdnCodClientes & ", "
			hdnCodClientes = hdnCodClientes & rs(0)
		rs.movenext
		loop
				
		hdnCodPaises = ""
		sql = "select CountryID from OLKBNOCRY where BannerID = " & Request("BannerID")
		set rs = conn.execute(sql)
		do while not rs.eof
			If hdnCodPaises <> "" Then hdnCodPaises = hdnCodPaises & ", "
			hdnCodPaises = hdnCodPaises & "'" & rs(0) & "'"
		rs.movenext
		loop
		
		hdnCodSecciones = ""
		sql = "select SecType + Convert(nvarchar(20),SecID) from OLKBNSections where BannerID = " & Request("BannerID")
		set rs = conn.execute(sql)
		do while not rs.eof
			If hdnCodSecciones <> "" Then hdnCodSecciones = hdnCodSecciones & ", "
			hdnCodSecciones = hdnCodSecciones& "'" & rs(0) & "'"
		rs.movenext
		loop
		
	Else
		GroupID = -1
		BannerDesc = ""
		Query = ""
	End If
	
 	sql = "select T0.GroupId, IsNull(AlterGroupName, GroupName) GroupName " & _
 			"from OLKBNGroups T0 " & _
 			"left outer join OLKBNGroupsAlterNames T1 on T1.GroupID = T0.GroupID and T1.LanID = " & Session("LanID") & " " & _
			 "where T0.groupID >= 0 "& _
			 "Order by T0.GroupId"
	set rs = conn.Execute(sql)
	
%>
<script language="javascript" src="funciones.js"></script>
<script language="javascript">
function selectOpciones(page, width, height)
{
	OpenWin = window.open(page+'&pop=Y','OpenWin', 'resizable=0,top=78,scrollbars=1, width='+width+',height='+height+'');
	OpenWin.moveTo((screen.width-320)/2,(screen.height-600)/2);  

}
function pasarToCode(strCod, opcion) 
{ 
	switch (opcion) {
   	case 'clientes':
    	document.frmBannerNew.hdnCodClientes.value = strCod;
    	break; 
   	case 'paises':
    	document.frmBannerNew.hdnCodPaises.value = strCod; 
    	break;
   	case 'secciones':
    	document.frmBannerNew.hdnCodSecciones.value = strCod;
    	break;   
	} 
}

function Validar()
{
	if (document.frmBannerNew.txtLink.value == '')
	{
		alert('<%=getadminBNEditLngStr("LtxtValBanLnk")%>');
		document.frmBannerNew.txtLink.focus();
		return false;
	}
  	if (document.frmBannerNew.txtPicture.value == "")
  	{ 
		alert('<%=getadminBNEditLngStr("LtxtValBanImg")%>');
  		return false; 
  	}
  	else if (document.frmBannerNew.valtxtQuery.value == 'Y' && document.frmBannerNew.txtQuery.value != '')
  	{
  		alert('<%=getadminBNEditLngStr("LtxtValBanQry")%>');
  		document.frmBannerNew.btnVerfyFilter.focus();
  		return false;
  	}
	return true;
}

var imgField
var imgImage
function uploadPic(Field, Img)
{
	imgField = Field;
	imgImage = Img;
	OpenWin = window.open('Upload/fileupload.aspx?ID=<%=Session("ID")%>&style=admin/style/style_pop.css&Source=Admin', 'OpenWin', 'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=yes,width=450, height=150' );
}
function changepic(filename) {
imgField.value = filename;
imgImage.src = "pic.aspx?dbName=<%=Session("olkdb")%>&maxSize=250&filename="+filename;
}
</script>
<table border="0" width="100%" id="table1" cellpadding="0">
	<form method="POST" action="adminSubmit.asp" onsubmit="javascript:return Validar();" name="frmBannerNew">
		<input type="hidden" name="BannerID" value="<%=Request("BannerID")%>">
		<input type="hidden" name="GroupId" value="<%=Request("GroupID")%>">
		<input type="hidden" name="cmd" value="saveBN">
		<% If Request("BannerID") = "" Then %>
		<input type="hidden" name="BannerDescTrad">
		<input type="hidden" name="BannerDescDef">
		<% End If %>
		<tr>
			<td bgcolor="#E1F3FD"><b><font face="Verdana" size="2">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><% If Request("BannerID") = "" Then %><%=getadminBNEditLngStr("DtxtAdd")%><% Else %><%=getadminBNEditLngStr("DtxtEdit")%><% End If %>&nbsp;<%=getadminBNEditLngStr("LtxtBanner")%></font></b></td>
		</tr>
		<tr>
			<td bgcolor="#F5FBFE">
			<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1">
			<font color="#4783C5"><%=getadminBNEditLngStr("LtxtAddEditDesc")%></font></font></td>
		</tr>
		<tr>
			<td>
			<table border="0" width="100%" id="table2" cellpadding="0">
				<tr>
					<td class="style1" style="width: 160px">
					<font face="Verdana" size="1"><strong><%=getadminBNEditLngStr("DtxtDescription")%></strong></font></td>
					<td class="style2">
					<table cellpadding="0" cellspacing="0" border="0">
						<tr>
							<td><input type="text" name="txtBannerDesc" size="50" value="<%=Server.HTMLEncode(BannerDesc)%>" maxlength="254" style="width: 308px"></td>
							<td><a href="javascript:doFldTrad('BN', 'BannerID', '<%=Request("BannerID")%>', 'AlterBannerDesc', 'T', <% If Request("BannerID") <> "" Then %>null<% Else %>document.frmBannerNew.BannerDescTrad<% End If %>);"><img src="images/trad.gif" alt="<%=getadminBNEditLngStr("DtxtTranslate")%>" border="0"></a></td>
						</tr>
					</table>
					</td>
				</tr>
				<tr>
					<td class="style1" style="width: 160px">
					<font face="Verdana" size="1"><strong><%=getadminBNEditLngStr("DtxtLink")%></strong></font></td>
					<td class="style2">
					<input type="text" name="txtLink" size="50" value="<%=Link%>" maxlength="254"></td>
				</tr>
				<tr>
					<td valign="top" class="style1" style="width: 160px">
					<font face="Verdana" size="1"><strong><%=getadminBNEditLngStr("LtxtImg")%></strong></font></td>
					<td class="style2">
					<table cellpadding="0" border="0">
						<tr>
							<td><img id="mainImg" src="pic.aspx?FileName=<% If IsNull(Picture) or Picture = "" Then %>n_a.gif<% Else %><%=Picture%><% End If %>&maxSize=250&dbName=<%=Session("olkdb")%>" border="1"></td>
							<td valign="bottom">&nbsp;<input type="button" value="<%=getadminBNEditLngStr("DtxtChange")%>" name="B2" class="OlkBtn" onclick="javascript:uploadPic(document.frmBannerNew.txtPicture, document.frmBannerNew.mainImg);"></td>
						</tr>
					</table>
					<input type="hidden" name="txtPicture" value="<%=Picture%>">
					</td>
				</tr>
				<tr>
					<td class="style1" style="width: 160px">
					<font face="Verdana" size="1"><strong><%=getadminBNEditLngStr("DtxtGroup")%></strong></font></td>
					<td class="style2"><select name="lstGroupID" size="1" class="input"><% 
				do while not rs.eof 
				%>
					<option value="<%=rs("GroupID")%>" <%if CInt(rs("groupid")) = GroupID then%>selected<%end if%>>
					<%=myHTMLEncode(rs("GroupName"))%></option>
					<% 
					rs.movenext
				loop
				set rs = nothing %></select> </td>
				</tr>
				<tr>
					<td class="style1" style="width: 160px">
					&nbsp;</td>
					<td class="style2">
					<input type="checkbox" class="noborder" <% If Status = "A" Then %>checked<% End If %> name="lstStatus" value="A" id="lstStatus"><label for="lstStatus"><font face="Verdana" size="1" color="#4783C5"><%=getadminBNEditLngStr("DtxtActive")%></font></label>
					</td>
				</tr>
				<tr>
					<td class="style1" style="width: 160px">
					<font face="Verdana" size="1"><strong><%=getadminBNEditLngStr("LtxtClientGrps")%></strong></font></td>
					<td class="style2">
					<input type="button" value="<%=getadminBNEditLngStr("DtxtEdit")%>" name="cmdClientes" onclick="javascript:selectOpciones('selectClientes.asp?hdnCodClientes=' + document.frmBannerNew.hdnCodClientes.value, 320,500)" class="OlkBtn"></td>
					<input type="hidden" name="hdnCodClientes" value="<%=hdnCodClientes%>">
				</tr>
				<tr>
					<td class="style1" style="width: 160px">
					<font face="Verdana" size="1"><strong><%=getadminBNEditLngStr("LtxtCountries")%></strong></font></td>
					<td class="style2">
					<input type="button" value="<%=getadminBNEditLngStr("DtxtEdit")%>" name="cmdPaises" onclick="javascript:selectOpciones('selectPaises.asp?hdnCodPaises=' + document.frmBannerNew.hdnCodPaises.value, 320,500)" class="OlkBtn">
					<input type="hidden" name="hdnCodPaises" value="<%=hdnCodPaises%>"></td>
				</tr>
				<tr>
					<td class="style1" style="width: 160px">
					<font face="Verdana" size="1"><strong><%=getadminBNEditLngStr("LtxtSections")%></strong></font></td>
					<td class="style2">
					<input type="button" value="<%=getadminBNEditLngStr("DtxtEdit")%>" name="cmdSecciones" onclick="javascript:selectOpciones('selectSecciones.asp?hdnCodSecciones=' + document.frmBannerNew.hdnCodSecciones.value, 320,500)" class="OlkBtn">
					<input type="hidden" name="hdnCodSecciones" value="<%=hdnCodSecciones%>"></td>
				</tr>
				<tr>
					<td valign="top" class="style1" style="width: 160px">
					<font face="Verdana" size="1" class="style4"><strong><%=getadminBNEditLngStr("DtxtQuery")%></strong><b><strong><br>
					</strong>
					<span class="style3">(<%=getadminBNEditLngStr("LtxtMustStart")%>&nbsp;<span dir="ltr"><strong><em>Select &#39;TRUE&#39;</em></strong></span><strong>)</strong></span></b></font></td>
					<td class="style2">
					<table cellpadding="0" cellspacing="0" border="0">
						<tr>
							<td rowspan="2">
								<textarea dir="ltr" rows="12" name="txtQuery" cols="77" onkeypress="javascript:document.frmBannerNew.btnVerfyFilter.src='images/btnValidate.gif';document.frmBannerNew.btnVerfyFilter.style.cursor = 'hand';document.frmBannerNew.valtxtQuery.value='Y';"><% If Not IsNull(Query) Then %><%=Server.HTMLEncode(Query)%><% End If %></textarea>
							</td>
							<td valign="top">
								<img style="cursor: hand; " src="images/qry_note.gif" id="btnNoteFilter" alt="<%=getadminBNEditLngStr("DtxtDefinition")%>" onclick="javascript:doFldNote(6, 'Query', '<%=Request("BannerID")%>', <% If Request("BannerID") <> "" Then %>null<% Else %>document.frmBannerNew.BannerDescDef<% End If %>);">
							</td>
						</tr>
						<tr>
							<td valign="bottom">
								<img src="images/btnValidateDis.gif" id="btnVerfyFilter" alt="<%=getadminBNEditLngStr("DtxtValidate")%>" onclick="javascript:if (document.frmBannerNew.valtxtQuery.value == 'Y')VerfyQuery();">
								<input type="hidden" name="valtxtQuery" value="N">
						</td>
						</tr>
					</table>
					</td>
				</tr>
				<tr>
					<td class="style1" style="width: 160px">
					<font face="Verdana" size="1"><strong><%=getadminBNEditLngStr("LtxtAvlVars")%></strong></font></td>
					<td class="style2">
					<font face="Verdana" size="1" color="#4783C5"> <span dir="ltr">@CardCode</span> = <%=getadminBNEditLngStr("LtxtCCode")%></font></td>
				</tr>
				<tr>
					<td class="style1" style="width: 160px">
					<font face="Verdana" size="1"><strong><%=getadminBNEditLngStr("LtxtStartDate")%></strong></font></td>
					<td class="style2">
					<table border="0" id="table6" cellspacing="0" cellpadding="0">
						<tr>
							<td><img border="0" src="images/cal.gif" id="btnStartDate" width="16" height="16" style="float:left;padding-left:1px;padding-top:1px"></td>
							<td>
							<input type="text" readonly name="txtStartDate" id="txtStartDate" onclick="btnStartDate.click();" size="15" value="<%=FormatDate(StartDate, False)%>"></td>
							<td><img border="0" src="images/remove.gif" width="16" height="16" onclick="javascript:document.frmBannerNew.txtStartDate.value='';"></td>
						</tr>
					</table>
					</td>
				</tr>
				<tr>
					<td class="style1" style="width: 160px">
					<font face="Verdana" size="1"><strong><%=getadminBNEditLngStr("LtxtEndDate")%></strong></font></td>
					<td class="style2">
					<table border="0" id="table7" cellspacing="0" cellpadding="0">
						<tr>
							<td><img border="0" src="images/cal.gif" id="btnEndDate" width="16" height="16" style="float:left;padding-left:1px;padding-top:1px"></td>
							<td>
					<input type="text" readonly name="txtEndDate" id="txtEndDate" size="15" onclick="btnEndDate.click();" value="<%=FormatDate(EndDate, False)%>"></td>
					<td><img border="0" src="images/remove.gif" width="16" height="16" onclick="javascript:document.frmBannerNew.txtEndDate.value='';"></td>
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
					<input type="submit" value="<%=getadminBNEditLngStr("DtxtApply")%>" name="btnApply" class="OlkBtn"></td>
					<td width="77">
					<input type="submit" value="<%=getadminBNEditLngStr("DtxtSave")%>" name="btnSave" class="OlkBtn"></td>
					<td><hr color="#0D85C6" size="1"></td>
					<td width="77">
					<input type="button" value="<%=getadminBNEditLngStr("DtxtCancel")%>" name="B1" class="OlkBtn" onclick="javascript:window.location.href='adminBN.asp';"></td>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td>&nbsp;</td>
		</tr>
		<input type="hidden" name="submitCmd" value="adminBN">
	</form>
</table>
<form name="frmVerfyQuery" action="verfyQuery.asp" method="post" target="iVerfyQuery">
	<iframe id="iVerfyQuery" name="iVerfyQuery" style="display: none" src=""></iframe>
	<input type="hidden" name="type" value="Banner">
	<input type="hidden" name="Query" value="">
	<input type="hidden" name="parent" value="Y">
</form>
<script type="text/javascript">
function VerfyQuery()
{
	document.frmVerfyQuery.Query.value = document.frmBannerNew.txtQuery.value;
	document.frmVerfyQuery.submit();
}
function VerfyQueryVerified()
{
	//document.frmBannerNew.btnVerfyQry.disabled = true;
	document.frmBannerNew.btnVerfyFilter.src='images/btnValidateDis.gif'
	document.frmBannerNew.btnVerfyFilter.style.cursor = '';
	document.frmBannerNew.valtxtQuery.value='N';
}

Calendar.setup({
    inputField     :    "txtStartDate",     // id of the input field
    ifFormat       :    "<%=GetCalendarFormatString%>",      // format of the input field
    button         :    "btnStartDate",  // trigger for the calendar (button ID)
    align          :    "Tl",           // alignment (defaults to "Bl")
    singleClick    :    true
});
Calendar.setup({
    inputField     :    "txtEndDate",     // id of the input field
    ifFormat       :    "<%=GetCalendarFormatString%>",      // format of the input field
    button         :    "btnEndDate",  // trigger for the calendar (button ID)
    align          :    "Tl",           // alignment (defaults to "Bl")
    singleClick    :    true
});
</script>

<!--#include file="bottom.asp" -->
<% addLngPathStr = "ventas/" %>
<!--#include file="lang/activityX.asp" -->
<%
set rs = Server.CreateObject("ADODB.recordset")
AsignedSlp = not myAut.HasAuthorization(60)

If Request("orden1") <> "" Then 
	orden1 = Request("orden1")
	orden2 = Request("orden2")
Else
	orden1 = myApp.ActOrdr1
	orden2 = myApp.ActOrdr2
End If

sourceType = Request("cmbSourceType")

sql = ""

If sourceType = "" or sourceType = "O" Then
	sql = sql & "select T0.LogNum TransNum, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OSLP', 'SlpName', T5.SlpCode, T5.SlpName) SlpName, " & _
			"OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OUSR', 'U_NAME', T6.INTERNAL_K, T6.U_Name) AttendUser, " & _
			"OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRG', 'GroupName', T2.GroupCode, GroupName) GroupName, " & _
			"T2.Country, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRY', 'Name', T1.Country, Name) CountryName, " & _
			"OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRD', 'CardName', T1.CardCode, T2.CardName) CardName, T1.CardCode collate database_default CardCode, " & _
			"IsNull(Details, '') collate database_default Comments, Recontact DocDate, Convert(int,Recontact) CntctDateSort, Closed collate database_default Closed, " & _
			"Inactive collate database_default Inactive, Action collate database_default Action, " & _
			"case when exists(select 'A' from ocrd where cardcode = T1.cardcode collate database_default) Then 'True' ELSE 'False' End Verfy, 'O' SourceType " & _
			"from R3_Obscommon..tlog T0 " & _
			"inner join r3_obscommon..TCLG T1 on T1.LogNum = T0.LogNum " & _
			"inner join ocrd T2 on T2.CardCode = T1.CardCode collate database_default " & _
			"left outer join ocry T3 on T3.code = T2.Country collate database_default " & _
			"inner join ocrg T4 on T4.GroupCode = T2.GroupCode " & _
			"left outer join oslp T5 on T5.slpcode = T1.SlpCode " & _
			"left outer join OUSR T6 on T6.INTERNAL_K = T1.AttendUser " & _
			"inner join R3_ObsCommon..TLOGControl X0 on X0.LogNum = T0.LogNum and X0.appId = 'TM-OLK' " & _
			"where Company = N'" & Session("olkdb") & "' and Object = 33 and T0.status = 'R' " & getActivityXFilter("O") & " "
End If

If sourceType = "" Then sql = sql & " union "

If sourceType = "" or sourceType = "S" Then
	sql = sql & "select T1.ClgCode TransNum, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OSLP', 'SlpName', T5.SlpCode, T5.SlpName) SlpName, " & _
			"OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OUSR', 'U_NAME', T6.INTERNAL_K, T6.U_Name) AttendUser, " & _
			"OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRG', 'GroupName', T2.GroupCode, GroupName) GroupName, " & _
			"T2.Country, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRY', 'Name', T1.Country, Name) CountryName, " & _
			"OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRD', 'CardName', T1.CardCode, T2.CardName) CardName, T1.CardCode, " & _
			"IsNull(Details, '') Comments, Recontact DocDate, Convert(int,Recontact) CntctDateSort, Closed, Inactive, Action, " & _
			"case when exists(select 'A' from ocrd where cardcode = T1.cardcode collate database_default) Then 'True' ELSE 'False' End Verfy, 'S' SourceType " & _
			"from OCLG T1 " & _
			"inner join ocrd T2 on T2.CardCode = T1.CardCode collate database_default " & _
			"left outer join ocry T3 on T3.code = T2.Country collate database_default " & _
			"inner join ocrg T4 on T4.GroupCode = T2.GroupCode " & _
			"left outer join OUSR T6 on T6.INTERNAL_K = T1.AttendUser " & _
			"left outer join OHEM T7 on T7.userId = T1.AttendUser " & _
			"left outer join oslp T5 on T5.slpcode = T7.salesPrson " & _
			"where T1.Closed = 'N' and T1.Inactive = 'N' " & getActivityXFilter("S") & " "
End If

sql = sql & "order by " & orden1 & " " & orden2
RS.CursorLocation = 3 ' adUseClient

rs.open sql, conn
RS.PageSize = 40
nPageCount = RS.PageCount
If Request("Page") <> "" Then nPage = CLng(Request("Page")) Else nPage = 1
If nPage < 1 Or nPage > nPageCount Then	nPage = 1
If Not Rs.Eof then RS.AbsolutePage = nPage
%>
<script language="javascript">
function listPendAlert() {
	alert('<%=getactivityXLngStr("LtxtDisObj")%>'.replace('{0}', "<%=getactivityXLngStr("DtxtActivity")%>"));
}
function valFrm()
{
	if (document.frmActivityX.chkDel.length)
	{
		var found = false;
		for (var i = 0;i<document.frmActivityX.chkDel.length;i++)
		{
			if (document.frmActivityX.chkDel[i].checked)
			{
				found = true;
				break;
			}
		}
		if (!found)
		{
			alert('<%=getactivityXLngStr("LtxtValSelAct")%>');
			return false;
		}
	}
	else
	{
		if (!document.frmActivityX.chkDel.checked)
		{
			alert('<%=getactivityXLngStr("LtxtValSelAct")%>');
			return false;
		}
	}
	return confirm('<%=getactivityXLngStr("LtxtConfDelAct")%>');
}
</script>
<div align="center">
<table border="0" cellpadding="0" width="100%">
<form name="frmActivityX" method="post" action="ventas/docdel.asp" onsubmit="javascript:return valFrm();">
	<tr class="GeneralTlt">
		<td><%=getactivityXLngStr("LttlPendActs")%></td>
	</tr>
	<% If rs.PageCount > 1 Then %>
	<tr>
		<td><% doActivityXPages %></td>
	</tr>
	<% End If %>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table2">
			<tr class="FirmTlt3">
				<td align="center" style="width: 15px">&nbsp;</td>
				<td align="center" style="width: 18px">&nbsp;</td>
				<td align="center" style="cursor: hand" onclick="javascript:doSort('T0.LogNum');" <% doVentasXSortBG("T0.LogNum")%>><%=getactivityXLngStr("DtxtLogNum")%><% doVentasXSortImg("T0.LogNum")%></td>
				<td align="center" style="cursor: hand" onclick="javascript:doSort('SlpName');" <% doVentasXSortBG("SlpName")%>><%=txtAgent%><% doVentasXSortImg("SlpName")%></td>
				<td align="center" style="cursor: hand" onclick="javascript:doSort('AttendUser');" <% doVentasXSortBG("AttendUser")%>><%=getactivityXLngStr("LtxtAsignedTo")%><% doVentasXSortImg("AttendUser")%></td>
				<% If strScriptName <> "activeclient.asp" Then %>
				<td align="center" style="cursor: hand" onclick="javascript:doSort('T1.CardCode');" <% doVentasXSortBG("T1.CardCode")%>><%=getactivityXLngStr("DtxtClient")%><% doVentasXSortImg("T1.CardCode")%></td>
				<td align="center" style="cursor: hand" onclick="javascript:doSort('CardName');" <% doVentasXSortBG("CardName")%>><%=getactivityXLngStr("DtxtName")%><% doVentasXSortImg("CardName")%></td>
				<% End If %>
				<td align="center" style="cursor: hand" onclick="javascript:doSort('GroupName');" <% doVentasXSortBG("GroupName")%>><%=getactivityXLngStr("DtxtGroup")%><% doVentasXSortImg("GroupName")%></td>
				<% If strScriptName <> "activeclient.asp" Then %><td align="center" style="cursor: hand; width: 30px;" onclick="javascript:doSort('Country');"><% doVentasXSortImg("Country")%></td><% End If %>
				<td align="center" style="cursor: hand; width: 75px;" onclick="javascript:doSort('CntctDateSort');" <% doVentasXSortBG("CntctDateSort")%>><%=getactivityXLngStr("DtxtDate")%><% doVentasXSortImg("CntctDateSort")%></td>
				<td align="center" style="cursor: hand" onclick="javascript:doSort('Action');" <% doVentasXSortBG("Action")%>><%=getactivityXLngStr("DtxtActivity")%><% doVentasXSortImg("Action")%></td>
			</tr>
		  <%  if not rs.eof then
		  do while not (rs.eof Or RS.AbsolutePage <> nPage )
		  Enable = myApp.EnableOCLG %>
			<tr class="GeneralTbl">
				<td style="width: 15px; height: 15px;" align="center">
				<img src="images/checkbox_off.jpg" border="0" onclick="doCheckDel(this, <%=rs("TransNum")%>);" <% If rs("SourceType") = "S" Then %>style="visibility: hidden; "<% End If %>>
				<input type="checkbox" name="chkDel" id="chkDel<%=rs("TransNum")%>" value="<%=rs("TransNum")%>" style="display: none;"></td>
				<td style="width: 18px; height: 15px;" align="center">
				<p align="center">
				<a href="javascript:<% If Enable Then %>doGoAct('<%=rs("SourceType")%>', '<%=rs("TransNum")%>');<% Else %>listPendAlert();<% End If %>">
				<img border="0" src="design/0/images/<%=Session("rtl")%>felcahSelect.gif" width="15" height="13"></a></td>
				<td style="height: 15px">
				<table cellpadding="0" cellspacing="0" border="0" style="width: 100%;">
              	<tr class="GeneralTbl">
              		<td><%=RS("TransNum")%></td>
              		<td style="width: 13px; "><img src="images/icon_activity_<%=rs("SourceType")%>.gif"></td>
              	</tr>
              	</table></td>
				<td style="height: 15px"><%=RS("SlpName")%>&nbsp;</td>
				<td style="height: 15px"><%=RS("AttendUser")%>&nbsp;</td>
				<% If strScriptName <> "activeclient.asp" Then %>
				<td style="height: 15px"><%=RS("CardCode")%>&nbsp;</td>
				<td style="height: 15px"><% If Not isNull(rs("CardName")) Then %><%=RS("cardname")%><% End If %>&nbsp;</td>
				<% End If %>
				<td style="height: 15px"><% If Not isNull(rs("GroupName")) Then %><%=RS("GroupName")%><% End If %>&nbsp;</td>
				<% If strScriptName <> "activeclient.asp" Then %><td style="width: 30px; text-align: center; height: 15px;">
				<img src="images/country/pic.aspx?filename=<%=rs("Country")%>.gif&MaxHeight=15" alt="<%=rs("CountryName")%>">
				</td><% End If %>
				<td style="text-align: center; width: 75px; height: 15px;"><%=FormatDate(RS("DocDate"), True)%>&nbsp;</td>
				<td style="height: 15px"><% Select Case rs("Action")
					Case "C"
						Response.Write getactivityXLngStr("DtxtConv")
					Case "M"
						Response.Write getactivityXLngStr("DtxtMeeting")
					Case "E"
						Response.Write getactivityXLngStr("DtxtNote")
					Case "O"
						Response.Write getactivityXLngStr("DtxtOther")
					Case "T"
						Response.Write getactivityXLngStr("DtxtTask")
					End Select %></td>
			</tr>
			<% If rs("Comments") <> "" Then %>
			<tr class="GeneralTbl">
				<td colspan="<% If strScriptName <> "activeclient.asp" Then %>11<% Else %>10<% End If %>">
				<%=getactivityXLngStr("DtxtObservations")%>:&nbsp;<%=rs("Comments")%></td>
			</tr>
			<% End If %>
			 <% rs.movenext
			 loop %>
			<tr class="GeneralTblBold2">
				<td colspan="<% If strScriptName <> "activeclient.asp" Then %>11<% Else %>10<% End If %>">
				<input type="submit" name="btnDel" value="<%=getactivityXLngStr("DtxtDelete")%>"><input type="hidden" name="go2" value="A"></td>
			</tr>
			<% Else %>
			<tr class="GeneralTblBold2">
				<td colspan="<% If strScriptName <> "activeclient.asp" Then %>11<% Else %>10<% End If %>">
				<p align="center"><%=getactivityXLngStr("DtxtNoData")%></td>
			</tr>
			<% End If %>
		</table>
		</td>
	</tr>	<% If rs.PageCount > 1 Then %>
	<tr>
		<td><% doActivityXPages %></td>
	</tr>
	<% End If %>
<% for each Item in Request.Form
	If Item <> "chkDel" and Item <> "btnDel" and Item <> "go2" and Item <> "orden1" and Item <> "orden2" Then %><input type="hidden" name="<%=Item%>" value="<%=Request.Form(Item)%>"><% End If
	next 
	
	for each Item in Request.QueryString
	If Item <> "chkDel" and Item <> "btnDel" and Item <> "go2" and Item <> "orden1" and Item <> "orden2" Then %><input type="hidden" name="<%=Item%>" value="<%=Request.QueryString(Item)%>"><% End If
	next %>
	<input type="hidden" name="orden1" value="<%=Request("orden1")%>">
	<input type="hidden" name="orden2" value="<%=Request("orden2")%>">
	</form>
</table>
</div>

<% 
Sub doActivityXPages %>
<table cellpadding="0" cellspacing="2" border="0" width="100%">
	<tr>
	<% If nPage > 1 Then %>
		<td width="15" class="FirmTlt3">
		<p align="center">
		<a href="javascript:goPage(<%= nPage - 1 %>);">
		<img border="0" src="design/0/images/<%=Session("rtl")%>prev_icon_trans.gif" width="15" height="15"></a></td>
		<% End If %>
		<td class="FirmTlt3">
		<p align="center" dir="ltr"><% if rs.PageCount > 1 then
        For I = 1 To rs.PageCount
			If I = nPage Then %>
				<font size="3">
				<b><%= I %></b></font>
				<% Else %>
				<a class="LnkSearchPaginacion" href="javascript:goPage(<%= I %>);"><%= I %></a>
				<% End If
		Next 'I
		end if %></td>
		<% If nPage < rs.PageCount Then %>
		<td width="15" class="FirmTlt3">
		<p align="center">
		<a href="javascript:goPage(<%= nPage + 1 %>);">
		<img border="0" src="design/0/images/<%=Session("rtl")%>next_icon_trans.gif" width="15" height="15"></a></td>
		<% End If %>
	</tr>
</table>
<% End Sub
Sub doVentasXSortImg(c)
	If LCase(orden1) = LCase(c) Then
		If orden2 = "asc" Then
			Response.Write "<img src=""images/arrow_up.gif"">"
		Else
			Response.Write "<img src=""images/arrow_down.gif"">"
		End If
	End If
End Sub 
Sub doVentasXSortBG(c)
	If LCase(orden1) = LCase(c) Then Response.Write "class=""GeneralTblBold2HighLight"""
End Sub %>
<script language="javascript">
function doCheckDel(Img, LogNum)
{
	if (!document.getElementById('chkDel' + LogNum).checked)
	{
		document.getElementById('chkDel' + LogNum).checked = true;
		Img.src = 'images/checkbox_on.jpg';
	}
	else
	{
		document.getElementById('chkDel' + LogNum).checked = false;
		Img.src = 'images/checkbox_off.jpg';
	}
}
function goPage(p) { document.frmGoX.page.value = p; document.frmGoX.submit(); }
function doSort(c)
{
	document.frmGoX.orden1.value = c;
	if ('<%=orden1%>' == c)
	{
		if ('<%=orden2%>' == 'asc')
			document.frmGoX.orden2.value = 'desc';
		else
			document.frmGoX.orden2.value = 'asc';
	}
	else
	{
		document.frmGoX.orden2.value = 'asc';
	}
	document.frmGoX.page.value = 1;
	document.frmGoX.submit();
}
function delAct(LogNum)
{
	if(!confirm('<%=getactivityXLngStr("LtxtConfDelAct")%>'.replace('{0}', LogNum))) return;
	doMyLink('ventas/docdel.asp', 'retval='+LogNum+varx, '');
}


function doGoAct(sourceType, transNum)
{
	switch (sourceType)
	{
		case 'O':
			document.doGoAct.LogNum.value = transNum;
			document.doGoAct.submit();
			break;
		case 'S':
			document.doGoEditAct.ClgCode.value = transNum;
			document.doGoEditAct.submit();
			break;
	}
}
</script>
<form name="doGoAct" method="post" action="addActivity/goActivity.asp">
<input type="hidden" name="LogNum" value="">
</form>
<form name="doGoEditAct" action="addActivity/goEditActivity.asp" method="post">
<input type="hidden" name="ClgCode" value="">
</form>
<form name="frmGoX" method="post" action="<%=strScriptName%>">
<input type="hidden" name="page" value="">
<input type="hidden" name="retval" value="">
<% 
varx = ""
for each Item in Request.Form 
	If Item <> "retval" Then
	varx = varx & "&" & Item & "=" & Request.Form(Item)
	If Item <> "page" Then %>
	<input type="hidden" name="<%=Item%>" value="<%=Request.Form(Item)%>">
<%	End If
	End If
next 

for each Item in Request.QueryString
	If Item <> "retval" Then
	varx = varx & "&" & Item & "=" & Request.QueryString(Item) 
	If Item <> "page" Then %>
	<input type="hidden" name="<%=Item%>" value="<%=Request.QueryString(Item)%>">
<%	End If
	End If
next %>
<% If Request("orden1") = "" and Request("page") = "" Then %>
<input type="hidden" name="orden1" value="">
<input type="hidden" name="orden2" value="">
<% End If %>
</form>
<script>
var varx = '<%=Replace(varx, "'", "\'")%>'
</script>
<%

Function getActivityXFilter(ByVal FilterType)

	Select Case FilterType
		Case "O"
			fldTrans = "T0.LogNum"
		Case "S"
			fldTrans = "T1.ClgCode"
	End Select
	
	cCode = ""
	If Request("CardCodeFrom") <> "" Then cCode = " and T1.CardCode >= N'" & saveHTMLDecode(Request("CardCodeFrom"), False) & "' "
	If Request("CardCodeTo") <> "" Then cCode = cCode & " and T1.CardCode <= N'" & saveHTMLDecode(Request("CardCodeTo"), False) & "' "
	If strScriptName = "activeclient.asp" Then cCode = " and T1.CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "' "
	
	LogNum = ""
	If Request("LogNumFrom") <> "" Then LogNum = " and " & fldTrans & " >= " & Request("LogNumFrom") & " "
	If Request("LogNumTo") <> "" Then LogNum = LogNum & " and " & fldTrans & " <= " & Request("LogNumTo") & " "
	
	'If Request("all") <> "Y" Then SlpCode = "and T0.SlpCode = " & Session("vendid")
	
	GroupCode = ""
	Country = ""
	
	If Request("GroupNameFrom") <> "" or Request("GroupNameTo") <> "" Then
		GroupCode = GroupCode & " and (( "
		
		If Request("GroupNameFrom") <> "" Then GroupCode = GroupCode & " T4.GroupName >= N'" & saveHTMLDecode(Request("GroupNameFrom"), False) & "' "
		If Request("GroupNameFrom") <> "" and Request("GroupNameTo") <> "" Then GroupCode = GroupCode & " and "
		If Request("GroupNameTo") <> "" Then GroupCode = GroupCode & " T4.GroupName <= N'" & saveHTMLDecode(Request("GroupNameTo"), False) & "' "
		
		GroupCode = GroupCode & ") or T2.GroupCode in (select PK " & _
						"	from OMLT X0 " & _
						"	inner join MLT1 X1 on X1.TranEntry = X0.TranEntry " & _
						"	where TableName = 'OCRG' and FieldAlias = 'GroupName' and LangCode = OLKCommon.dbo.OLKGetSBOLang(" & Session("LanID") & ") "
						
		If Request("GroupNameFrom") <> "" Then GroupCode = GroupCode & " and Convert(nvarchar(100),Trans) >= N'" & saveHTMLDecode(Request("GroupNameFrom"), False) & "' "
		If Request("GroupNameTo") <> "" Then GroupCode = GroupCode & " and Convert(nvarchar(100),Trans) <= N'" & saveHTMLDecode(Request("GroupNameTo"), False) & "' "
		
		GroupCode = GroupCode & ") ) "
	End If
	
	If Request("CountryFrom") <> "" or Request("CountryTo") <> "" Then
		Country = Country & " and (( "
		
		If Request("CountryFrom") <> "" Then Country = Country & " T3.Name >= N'" & saveHTMLDecode(Request("CountryFrom"), False) & "' "
		If Request("CountryFrom") <> "" and Request("CountryTo") <> "" Then Country = Country & " and "
		If Request("CountryTo") <> "" Then Country = Country & " T3.Name <= N'" & saveHTMLDecode(Request("CountryTo"), False) & "' "
		
		Country = Country & ") or T3.Code in (select PK " & _
						"	from OMLT X0 " & _
						"	inner join MLT1 X1 on X1.TranEntry = X0.TranEntry " & _
						"	where TableName = 'OCRY' and FieldAlias = 'Name' and LangCode = OLKCommon.dbo.OLKGetSBOLang(" & Session("LanID") & ") "
						
		If Request("CountryFrom") <> "" Then Country = Country & " and Convert(nvarchar(100),Trans) >= N'" & saveHTMLDecode(Request("CountryFrom"), False) & "' "
		If Request("CountryTo") <> "" Then Country = Country & " and Convert(nvarchar(100),Trans) <= N'" & saveHTMLDecode(Request("CountryTo"), False) & "' "
		
		Country = Country & ") ) "
	End If
	
	If Request("CardString") <> "" Then CardString = "and (T2.CardCode like N'%" & saveHTMLDecode(Request("CardString"), False) & "%' or T1.CardName like N'%" & saveHTMLDecode(Request("CardString"), False) & "%')"
	If Request("DocType") <> "" Then 
		DocType = "and Action = '" & Request("DocType") & "'"
	End If
	
	SlpFilter = ""
	If Request("SlpCodeFrom") <> "" Then SlpFilter = " and OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OSLP', 'SlpName', T5.SlpCode, T5.SlpName) collate database_default >= N'" & saveHTMLDecode(Request("SlpCodeFrom"), False) & "' "
	If Request("SlpCodeTo") <> "" Then SlpFilter = SlpFilter & " and OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OSLP', 'SlpName', T5.SlpCode, T5.SlpName) collate database_default <= N'" & saveHTMLDecode(Request("SlpCodeTo"), False) & "' "

	If Request("AttendUserFrom") <> "" Then SlpFilter = " and OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OUSR', 'U_NAME', T6.INTERNAL_K, T6.U_Name) collate database_default >= N'" & saveHTMLDecode(Request("AttendUserFrom"), False) & "' "
	If Request("AttendUserTo") <> "" Then SlpFilter = SlpFilter & " and OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OUSR', 'U_NAME', T6.INTERNAL_K, T6.U_Name) collate database_default <= N'" & saveHTMLDecode(Request("AttendUserTo"), False) & "' "
	
	If Request("dtFrom") <> "" Then DocDate = " and DateDiff(day,Convert(datetime,'" & SaveSqlDate(Request("dtFrom")) & "',120), Recontact) >= 0 "
	If Request("dtTo") <> "" Then DocDate = DocDate & " and DateDiff(day,Recontact, Convert(datetime,'" & SaveSqlDate(Request("dtTo")) & "',120)) >= 0"
	
	If AsignedSlp or not myAut.HasAuthorization(97) Then 
		Select Case FilterType
			Case "O"
				SlpCode1 = "and T1.SlpCode = " & Session("VendId") & " "
			Case "S"
				SlpCode1 = "and T7.salesPrson = " & Session("VendId") & " "
		End Select
	End If
	
	retVal = SlpCode1 & cCode & LogNum & DocDate & CardString & Country & GroupCode & DocType & DocTypeAdd & SlpFilter
	
	If Not IsNull(myApp.AgentClientsFilter) Then
		retVal = retVal & " and T1.CardCode collate database_default not in (" & Replace(Replace(myApp.AgentClientsFilter, "@SlpCode", Session("vendid")), "@Type", 1) & ") "
		retVal = retVal & " and T1.CardCode collate database_default not in (" & Replace(Replace(myApp.AgentClientsFilter, "@SlpCode", Session("vendid")), "@Type", 5) & ") "
	End If
	
	getActivityXFilter = retVal
End Function 
%>
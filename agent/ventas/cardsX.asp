<% addLngPathStr = "ventas/" %>
<!--#include file="lang/cardsX.asp" -->


<%
set rs = Server.CreateObject("ADODB.RecordSet")
sqll = _
"select T0.LogNum, IsNull(T0.CardCode, '') CardCode, IsNull(T0.CardName, '') CardName, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRG', 'GroupName', T0.GroupCode, GroupName) GroupName, CardType, " & _
"IsNull(T3.Name, '') Country, T1.SubDate CreateDate, Convert(int,T1.SubDate) CreateDateSort, Status, " & _
"Case Status When 'R' Then '" & getcardsXLngStr("DtxtPend") & "' When 'H' Then '" & getcardsXLngStr("DtxtConfirmed") & "' End StatusStr, T1.ErrMessage " & _
"from R3_ObsCommon..TCRD T0 " & _
"inner join R3_ObsCommon..TLOG T1 on T1.LogNum = T0.LogNum " & _
"left outer join OCRG T2 on T2.GroupCode = T0.GroupCode " & _
"left outer join OCRY T3 on T3.Code = T0.Country collate database_default " & _
"inner join R3_ObsCommon..TLOGControl L0 on L0.LogNum = T0.LogNum and L0.AppID = 'TM-OLK' " & _
"where T1.Company = db_name() and Status in ('R', 'H') "

If Request("LogNumFrom") <> "" Then sqll = sqll & " and T0.LogNum >= " & Request("LogNumFrom") & " "
If Request("LogNumTo") <> "" Then sqll = sqll & " and T0.LogNum <= " & Request("LogNumTo") & " "
If Request("dtFrom") <> "" Then sqll = sqll & " and DateDiff(day,Convert(datetime,'" & SaveSqlDate(Request("dtFrom")) & "',120), T1.SubDate) >= 0 "
If Request("dtTo") <> "" Then sqll = sqll & " and DateDiff(day,T1.SubDate, Convert(datetime,'" & SaveSqlDate(Request("dtTo")) & "',120)) >= 0"
If Request("CardCodeFrom") <> "" Then sqll = sqll & " and T0.CardCode >= N'" & saveHTMLDecode(Request("CardCodeFrom"), False) & "' "
If Request("CardCodeTo") <> "" Then sqll = sqll & " and T0.CardCode <= N'" & saveHTMLDecode(Request("CardCodeTo"), False) & "' "


If Request("GroupNameFrom") <> "" or Request("GroupNameTo") <> "" Then
	sqll = sqll & " and (( "
	
	If Request("GroupNameFrom") <> "" Then sqll = sqll & " T2.GroupName >= N'" & saveHTMLDecode(Request("GroupNameFrom"), False) & "' "
	If Request("GroupNameFrom") <> "" and Request("GroupNameTo") <> "" Then sqll = sqll & " and "
	If Request("GroupNameTo") <> "" Then sqll = sqll & " T2.GroupName <= N'" & saveHTMLDecode(Request("GroupNameTo"), False) & "' "
	
	sqll = sqll & ") or T2.GroupCode in (select PK " & _
					"	from OMLT X0 " & _
					"	inner join MLT1 X1 on X1.TranEntry = X0.TranEntry " & _
					"	where TableName = 'OCRG' and FieldAlias = 'GroupName' and LangCode = OLKCommon.dbo.OLKGetSBOLang(" & Session("LanID") & ") "
					
	If Request("GroupNameFrom") <> "" Then sqll = sqll & " and Convert(nvarchar(100),Trans) >= N'" & saveHTMLDecode(Request("GroupNameFrom"), False) & "' "
	If Request("GroupNameTo") <> "" Then sqll = sqll & " and Convert(nvarchar(100),Trans) <= N'" & saveHTMLDecode(Request("GroupNameTo"), False) & "' "
	
	sqll = sqll & ") ) "
End If

If Request("CountryFrom") <> "" Then sqll = sqll & " and T3.Name >= N'" & saveHTMLDecode(Request("CountryFrom"), False) & "' "
If Request("CountryTo") <> "" Then sqll = sqll & " and T3.Name <= N'" & saveHTMLDecode(Request("CountryTo"), False) & "' "

If Request("CardType") <> "" Then
	sqll = sqll & " and T0.CardType = '" & Request("CardType") & "' "
Else
	CardType = ""
	If myAut.HasAuthorization(45) Then CardType = "'C'"
	If myAut.HasAuthorization(78) Then 
		If CardType <> "" Then CardType = CardType & ", "
		CardType = CardType & "'S'"
	End If
	If myAut.HasAuthorization(77) Then 
		If CardType <> "" Then CardType = CardType & ", "
		CardType = CardType & "'L'"
	End If
	sqll = sqll & " and T0.CardType in (" & CardType & ") "
End If

sqll = sqll & " order by " & Request("orden1") & " " & Request("orden2")

RS.CursorLocation = 3 ' adUseClient
rs.open sqll, conn
RS.PageSize = 40
nPageCount = RS.PageCount
nPage = CLng(Request.Form("Page"))
If nPage < 1 Or nPage > nPageCount Then	nPage = 1
If Not Rs.Eof then RS.AbsolutePage = nPage
%>
<script language="javascript">
function confReOpen(var1)
{
	return confirm('<%=getcardsXLngStr("LtxtConfReOpen")%>');
}
function valFrm()
{
	if (document.frmCardsX.chkDel.length)
	{
		var found = false;
		for (var i = 0;i<document.frmCardsX.chkDel.length;i++)
		{
			if (document.frmCardsX.chkDel[i].checked)
			{
				found = true;
				break;
			}
		}
		if (!found)
		{
			alert('<%=getcardsXLngStr("LtxtValSelCrd")%>');
			return false;
		}
	}
	else
	{
		if (!document.frmCardsX.chkDel.checked)
		{
			alert('<%=getcardsXLngStr("LtxtValSelCrd")%>');
			return false;
		}
	}
	return confirm('<%=getcardsXLngStr("LtxtConfDel")%>');
}

</script>
<div align="center">
<form name="frmCardsX" method="post" action="ventas/docdel.asp" onsubmit="javascript:return valFrm();">

<table border="0" cellpadding="0" width="100%" id="table1">
	<tr class="GeneralTlt">
		<td><%=getcardsXLngStr("LttlPendClient")%></td>
	</tr>
	<% If rs.PageCount > 1 Then %>
	<tr>
		<td>
		<% doCardXPages %>
		</td>
	</tr>
	<% End If %>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table2">
			<tr class="FirmTlt3">
				<td align="center" style="width: 15px">&nbsp;</td>
				<td align="center" style="cursor: hand; width: 15px;" onclick="javascript:doSort('T0.CardType');">&nbsp;<% doCardsXSortImg("T0.CardType")%></td>
				<td align="center" colspan="2" style="width: 100px; cursor: hand" onclick="javascript:doSort('T0.LogNum');" <% doCardsXSortBG("T0.LogNum")%>><%=getcardsXLngStr("DtxtLogNum")%><% doCardsXSortImg("T0.LogNum")%></td>
				<td align="center" style="cursor: hand" onclick="javascript:doSort('T0.CardCode');" <% doCardsXSortBG("T0.CardCode")%>><%=getcardsXLngStr("DtxtCode")%><% doCardsXSortImg("T0.CardCode")%></td>
				<td align="center" style="cursor: hand" onclick="javascript:doSort('T0.CardName');" <% doCardsXSortBG("T0.CardName")%>><%=getcardsXLngStr("DtxtName")%><% doCardsXSortImg("T0.CardName")%></td>
				<td align="center" style="cursor: hand" onclick="javascript:doSort('T2.GroupName');" <% doCardsXSortBG("T2.GroupName")%>><%=getcardsXLngStr("DtxtGroup")%><% doCardsXSortImg("T2.GroupName")%></td>
				<td align="center" style="cursor: hand" onclick="javascript:doSort('T3.Name');" <% doCardsXSortBG("T3.Name")%>><%=getcardsXLngStr("DtxtCountry")%><% doCardsXSortImg("T3.Name")%></td>
				<td align="center" style="cursor: hand" onclick="javascript:doSort('CreateDateSort');" <% doCardsXSortBG("CreateDateSort")%>><%=getcardsXLngStr("DtxtDate")%><% doCardsXSortImg("CreateDateSort")%></td>
				<td align="center" style="cursor: hand" onclick="javascript:doSort('Status');" <% doCardsXSortBG("Status")%>><%=getcardsXLngStr("DtxtState")%><% doCardsXSortImg("Status")%></td>
			</tr>
			<%  if not rs.eof then
 			do while not (rs.eof Or RS.AbsolutePage <> nPage )%>
			<tr class="GeneralTbl">
				<td style="width: 15px" align="center">
				<img src="images/checkbox_off.jpg" border="0" onclick="doCheckDel(this, <%=rs("LogNum")%>);">
				<input type="checkbox" name="chkDel" id="chkDel<%=rs("LogNum")%>" value="<%=rs("LogNum")%>" style="display: none;"></td>
				<td style="width: 15px"><b>
				<p align="center">
				<% Select Case rs("CardType")
					Case "C" %>
				<img src="ventas/images/icon_supplier.gif" alt="<%=txtClient%>">
				<%	Case "L" %>
				<img src="ventas/images/icon_lead.gif" alt="<%=getcardsXLngStr("DtxtLead")%>">
				<% Case "S" %>
				<img src="ventas/images/icon_client.gif" alt="<%=getcardsXLngStr("DtxtSupplier")%>">
				<% End Select%></b></td>
				<td width="15">
				<a href="javascript:<% If rs("Status") = "H" Then %>if(confReOpen())<% End If %>window.location.href='ventas/goCard.asp?LogNum=<%=RS("LogNum")%>'"><img border="0" src="design/0/images/<%=Session("rtl")%>felcahSelect.gif" width="15" height="13"></a></td>
				<td style="width: 85px;"><%=RS("LogNum")%></td>
				<td><%=RS("CardCode")%>&nbsp;</td>
				<td><%=RS("CardName")%>&nbsp;</td>
				<td><%=RS("GroupName")%>&nbsp;</td>
				<td><%=RS("Country")%>&nbsp;</td>
				<td align="center"><%=FormatDate(RS("CreateDate"), True)%>&nbsp;</td>
				<td align="center" <% If Not IsNull(rs("ErrMessage")) Then %>class="GeneralTblBold2HighLight"<% End If %>><%=rs("StatusStr")%>&nbsp;
				</td>
			</tr>
			<% If Not IsNull(rs("ErrMessage")) Then %>
			<tr class="GeneralTblBold2HighLight">
				<td colspan="10">
				<p align="center"><%=rs("ErrMessage")%></td>
			</tr>
			<% End If %>
			 <% rs.movenext
			 loop %>
			<tr class="GeneralTblBold2">
				<td colspan="10">
				<input type="submit" name="btnDel" value="<%=getcardsXLngStr("DtxtDelete")%>"><input type="hidden" name="go2" value="C"></td>
			</tr>
			<% Else %>
			<tr class="GeneralTbl">
				<td colspan="10">
				<p align="center"><%=getcardsXLngStr("DtxtNoData")%></td>
			</tr>
			<% End If %>
		</table>
		</td>
	</tr>
	<% If rs.PageCount > 1 Then %>
	<tr>
		<td>
		<% doCardXPages %>
		</td>
	</tr>
	<% End If %>
</table>
<% for each Item in Request.Form
If Item <> "chkDel" and Item <> "btnDel" and Item <> "go2" Then %><input type="hidden" name="<%=Item%>" value="<%=Request.Form(Item)%>"><% End If
next 

for each Item in Request.QueryString %><input type="hidden" name="<%=Item%>" value="<%=Request.QueryString(Item)%>"><% next %>
</form>
</div>

<% Sub doCardXPages %>
<table cellpadding="0" border="0" style="width: 100%;">
	<tr class="FirmTlt3">
		<% If nPage > 1 Then %><td width="15">
		<p align="center">
		<a href="javascript:goPage(<%= nPage - 1 %>);">
		<img border="0" src="design/0/images/<%=Session("rtl")%>prev_icon_trans.gif" width="15" height="15"></a></td><% End If %>
		<td>
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
		<td width="15">
		<p align="center">
		<a href="javascript:goPage(<%= nPage + 1 %>);">
		<img border="0" src="design/0/images/<%=Session("rtl")%>next_icon_trans.gif" width="15" height="15"></a></td><% End If %>
	</tr>
</table>
<% End Sub %>
<% 
Sub doCardsXSortImg(c)
	If LCase(Request("orden1")) = LCase(c) Then
		If Request("orden2") = "asc" Then
			Response.Write "<img src=""images/arrow_up.gif"">"
		Else
			Response.Write "<img src=""images/arrow_down.gif"">"
		End If
	End If
End Sub 
Sub doCardsXSortBG(c)
	If LCase(Request("orden1")) = LCase(c) Then Response.Write "class=""GeneralTblBold2HighLight"""
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
	if ('<%=Request("orden1")%>' == c)
	{
		if ('<%=Request("orden2")%>' == 'asc')
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
</script>
<form name="frmGoX" method="post" action="searchOpenedCards.asp">
<input type="hidden" name="page" value="">
<% 
varx = ""
for each Item in Request.Form 
	varx = varx & "&" & Item & "=" & Request.Form(Item)
	If Item <> "page" Then %>
<input type="hidden" name="<%=Item%>" value="<%=Request.Form(Item)%>">
<%	End If
next 
for each Item in Request.QueryString 
	varx = varx & "&" & Item & "=" & Request.QueryString (Item)
	If Item <> "page" Then %>
<input type="hidden" name="<%=Item%>" value="<%=Request.QueryString (Item)%>">
<%	End If
next  %>
</form>
<script>
var varx = '<%=varx%>'
</script>
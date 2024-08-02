<% addLngPathStr = "ventas/" %>
<!--#include file="lang/itemsX.asp" -->
<%
set rs = Server.CreateObject("ADODB.recordset")

sqll = _
"select T0.LogNum, IsNull(T0.ItemCode, '') ItemCode, IsNull(T0.ItemName, '') ItemName, " & _
"OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITB', 'ItmsGrpNam', T0.ItmsGrpCod, ItmsGrpNam) ItmsGrpNam, " & _
"OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OMRC', 'FirmName', T0.FirmCode, FirmName) FirmName, " & _
"T1.SubDate CreateDate, Convert(int,T1.SubDate) CreateDateSort, Status, " & _
"Case Status When 'R' Then '" & getitemsXLngStr("DtxtPend") & "' When 'H' Then '" & getitemsXLngStr("DtxtConfirmed") & "' End StatusStr, T1.ErrMessage " & _
"from R3_ObsCommon..TITM T0 " & _
"inner join R3_ObsCommon..TLOG T1 on T1.LogNum = T0.LogNum " & _
"left outer join OITB T2 on T2.ItmsGrpCod = T0.ItmsGrpCod " & _
"left outer join OMRC T3 on T3.FirmCode = T0.FirmCode " & _
"where T1.Company = db_name() and Status in ('R', 'H') "

If Request("LogNumFrom") <> "" Then sqll = sqll & " and T0.LogNum >= " & Request("LogNumFrom") & " "
If Request("LogNumTo") <> "" Then sqll = sqll & " and T0.LogNum <= " & Request("LogNumTo") & " "
If Request("dtFrom") <> "" Then sqll = sqll & " and DateDiff(day,Convert(datetime,'" & SaveSqlDate(Request("dtFrom")) & "',120), T1.SubDate) >= 0 "
If Request("dtTo") <> "" Then sqll = sqll & " and DateDiff(day,T1.SubDate, Convert(datetime,'" & SaveSqlDate(Request("dtTo")) & "',120)) >= 0"
If Request("ItemCodeFrom") <> "" Then sqll = sqll & " and T0.ItemCode >= N'" & saveHTMLDecode(Request("ItemCodeFrom"), False) & "' "
If Request("ItemCodeTo") <> "" Then sqll = sqll & " and T0.ItemCode <= N'" & saveHTMLDecode(Request("ItemCodeTo"), False) & "' "


If Request("ItmsGrpNamFrom") <> "" or Request("ItmsGrpNamTo") <> "" Then
	sqll = sqll & " and (( "
	
	If Request("ItmsGrpNamFrom") <> "" Then sqll = sqll & " T2.ItmsGrpNam >= N'" & saveHTMLDecode(Request("ItmsGrpNamFrom"), False) & "' "
	If Request("ItmsGrpNamFrom") <> "" and Request("ItmsGrpNamTo") <> "" Then sqll = sqll & " and "
	If Request("ItmsGrpNamTo") <> "" Then sqll = sqll & " T2.ItmsGrpNam <= N'" & saveHTMLDecode(Request("ItmsGrpNamTo"), False) & "' "
	
	sqll = sqll & ") or T2.ItmsGrpCod in (select PK " & _
					"	from OMLT X0 " & _
					"	inner join MLT1 X1 on X1.TranEntry = X0.TranEntry " & _
					"	where TableName = 'OITB' and FieldAlias = 'ItmsGrpNam' and LangCode = OLKCommon.dbo.OLKGetSBOLang(" & Session("LanID") & ") "
					
	If Request("ItmsGrpNamFrom") <> "" Then sqll = sqll & " and Convert(nvarchar(100),Trans) >= N'" & saveHTMLDecode(Request("ItmsGrpNamFrom"), False) & "' "
	If Request("ItmsGrpNamTo") <> "" Then sqll = sqll & " and Convert(nvarchar(100),Trans) <= N'" & saveHTMLDecode(Request("ItmsGrpNamTo"), False) & "' "
	
	sqll = sqll & ") ) "
End If
If Request("FirmNameFrom") <> "" or Request("FirmNameTo") <> "" Then
	sqll = sqll & " and (( "
	
	If Request("FirmNameFrom") <> "" Then sqll = sqll & " T3.FirmName >= N'" & saveHTMLDecode(Request("FirmNameFrom"), False) & "' "
	If Request("FirmNameFrom") <> "" and Request("FirmNameTo") <> "" Then sqll = sqll & " and "
	If Request("FirmNameTo") <> "" Then sqll = sqll & " T3.FirmName <= N'" & saveHTMLDecode(Request("FirmNameTo"), False) & "' "
	
	sqll = sqll & ") or T3.FirmCode in (select PK " & _
					"	from OMLT X0 " & _
					"	inner join MLT1 X1 on X1.TranEntry = X0.TranEntry " & _
					"	where TableName = 'OMRC' and FieldAlias = 'FirmName' and LangCode = OLKCommon.dbo.OLKGetSBOLang(" & Session("LanID") & ") "
					
	If Request("FirmNameFrom") <> "" Then sqll = sqll & " and Convert(nvarchar(100),Trans) >= N'" & saveHTMLDecode(Request("FirmNameFrom"), False) & "' "
	If Request("FirmNameTo") <> "" Then sqll = sqll & " and Convert(nvarchar(100),Trans) <= N'" & saveHTMLDecode(Request("FirmNameTo"), False) & "' "
	
	sqll = sqll & ") ) "
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
function confReOpen()
{
	return confirm('<%=getitemsXLngStr("LtxtConfReOpen")%>');
}
function valFrm()
{
	if (document.frmItemsX.chkDel.length)
	{
		var found = false;
		for (var i = 0;i<document.frmItemsX.chkDel.length;i++)
		{
			if (document.frmItemsX.chkDel[i].checked)
			{
				found = true;
				break;
			}
		}
		if (!found)
		{
			alert('<%=getitemsXLngStr("LtxtValSelItm")%>');
			return false;
		}
	}
	else
	{
		if (!document.frmItemsX.chkDel.checked)
		{
			alert('<%=getitemsXLngStr("LtxtValSelItm")%>');
			return false;
		}
	}
	return confirm('<%=getitemsXLngStr("LtxtConfDelItm")%>');
}
</script>
<form name="frmItemsX" method="post" action="ventas/docdel.asp" onsubmit="javascript:return valFrm();">

<table border="0" cellpadding="0" width="100%" id="table1">
	<tr class="GeneralTlt">
		<td><%=getitemsXLngStr("LttlPendItms")%></td>
	</tr>
	<% If rs.PageCount > 1 Then %>
	<tr>
		<td><% doItemsXPages %></td>
	</tr>
	<% End If %>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table2">
			<tr class="FirmTlt3">
				<td align="center" style="width: 15px">&nbsp;</td>
				<td align="center" style="width: 18px">&nbsp;</td>
				<td align="center" style="cursor: hand" onclick="javascript:doSort('T0.LogNum');" <% doItemsXSortBG("T0.LogNum")%>><%=getitemsXLngStr("DtxtLogNum")%><% doItemsXSortImg("T0.LogNum")%></td>
				<td align="center" style="cursor: hand" onclick="javascript:doSort('T0.ItemCode');" <% doItemsXSortBG("T0.ItemCode")%>><%=getitemsXLngStr("DtxtCode")%><% doItemsXSortImg("T0.ItemCode")%></td>
				<td align="center" style="cursor: hand" onclick="javascript:doSort('T0.ItemName');" <% doItemsXSortBG("T0.ItemName")%>><%=getitemsXLngStr("DtxtDescription")%><% doItemsXSortImg("T0.ItemName")%></td>
				<td align="center" style="cursor: hand" onclick="javascript:doSort('T2.ItmsGrpNam');" <% doItemsXSortBG("T2.ItmsGrpNam")%>><%=Server.HTMLEncode(txtAlterGrp)%><% doItemsXSortImg("T2.ItmsGrpNam")%></td>
				<td align="center" style="cursor: hand" onclick="javascript:doSort('T3.FirmName');" <% doItemsXSortBG("T3.FirmName")%>><%=Server.HTMLEncode(txtAlterFrm)%><% doItemsXSortImg("T3.FirmName")%></td>
				<td align="center" style="cursor: hand" onclick="javascript:doSort('CreateDateSort');" <% doItemsXSortBG("CreateDateSort")%>><%=getitemsXLngStr("DtxtDate")%><% doItemsXSortImg("CreateDateSort")%></td>
				<td align="center" style="cursor: hand" onclick="javascript:doSort('StatusStr');" <% doItemsXSortBG("StatusStr")%>><%=getitemsXLngStr("DtxtState")%><% doItemsXSortImg("StatusStr")%></td>
			</tr>
			  <%  if not rs.eof then
			  do while not (rs.eof Or RS.AbsolutePage <> nPage )%>
			<tr class="GeneralTbl">
				<td style="width: 15px;" align="center">
				<img src="images/checkbox_off.jpg" border="0" onclick="doCheckDel(this, <%=rs("LogNum")%>);">
				<input type="checkbox" name="chkDel" id="chkDel<%=rs("LogNum")%>" value="<%=rs("LogNum")%>" style="display: none;"></td>
				<td style="width: 18px;">
				<p align="center">
				<a href="javascript:<% If rs("Status") = "H" Then %>if(confReOpen())<% End If %>window.location.href='ventas/goItem.asp?LogNum=<%=RS("LogNum")%>'">
				<img border="0" src="design/0/images/<%=Session("rtl")%>felcahSelect.gif" width="15" height="13"></a></td>
				<td><%=RS("LogNum")%>&nbsp;</td>
				<td><%=RS("ItemCode")%>&nbsp;</td>
				<td><%=RS("ItemName")%>&nbsp;</td>
				<td><%=RS("ItmsGrpNam")%>&nbsp;</td>
				<td><%=RS("FirmName")%>&nbsp;</td>
				<td align="center"><%=FormatDate(RS("CreateDate"), False)%>&nbsp;</td>
				<td align="center" <% If Not IsNull(rs("ErrMessage")) Then %>class="GeneralTblBold2HighLight"<% End If %>><%=rs("StatusStr")%>&nbsp;</td>
			</tr>
			<% If Not IsNull(rs("ErrMessage")) Then %>
			<tr class="GeneralTblBold2HighLight">
				<td colspan="12">
				<p align="center"><%=rs("ErrMessage")%></td>
			</tr>
			<% End If %>
			 <% rs.movenext
			 loop %>
			<tr class="GeneralTblBold2">
				<td colspan="12">
				<input type="submit" name="btnDel" value="<%=getitemsXLngStr("DtxtDelete")%>"><input type="hidden" name="go2" value="I"></td>
			</tr>
			<% Else %>
			<tr class="GeneralTbl">
				<td colspan="9">
				<p align="center"><%=getitemsXLngStr("DtxtNoData")%></td>
			</tr>
			<% End If %>
		</table>
		</td>
	</tr>
	<% If rs.PageCount > 1 Then %>
	<tr>
		<td><% doItemsXPages %></td>
	</tr>
	<% End If %>
</table>
<% for each Item in Request.Form
If Item <> "chkDel" and Item <> "btnDel" and Item <> "go2" Then %><input type="hidden" name="<%=Item%>" value="<%=Request.Form(Item)%>"><% End If
next 

for each Item in Request.QueryString %><input type="hidden" name="<%=Item%>" value="<%=Request.QueryString(Item)%>"><% next %>

</form>
<% Sub doItemsXPages %>
<table cellpadding="0" border="0" style="width: 100%;">
	<tr class="FirmTlt3">
		<% If nPage > 1 Then %><td width="15">
		<p align="center">
		<a href="javascript:goPage(<%= nPage - 1 %>);"><img border="0" src="design/0/images/<%=Session("rtl")%>prev_icon_trans.gif" width="15" height="15"></a></td><% End If %>
		<td><p align="center" dir="ltr"><%  
		if rs.PageCount > 1 then
		For I = 1 To rs.PageCount
		If I = nPage Then %><b><font size="3"><%= I %></font></b>
		<% Else %><a class="LnkSearchPaginacion" href="javascript:goPage(<%= I %>);"><%= I %></a>
		<% End If
		Next 'I
		End If %></td>
		<% If nPage < rs.PageCount Then %>
		<td width="15">
		<p align="center">
		<a href="javascript:goPage(<%= nPage + 1 %>);"><img border="0" src="design/0/images/<%=Session("rtl")%>next_icon_trans.gif" width="15" height="15"></a></td><% End If %>
	</tr>
</table>
<% End Sub %>

<% 
Sub doItemsXSortImg(c)
	If LCase(Request("orden1")) = LCase(c) Then
		If Request("orden2") = "asc" Then
			Response.Write "<img src=""images/arrow_up.gif"">"
		Else
			Response.Write "<img src=""images/arrow_down.gif"">"
		End If
	End If
End Sub 
Sub doItemsXSortBG(c)
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
<form name="frmGoX" method="post" action="searchOpenedItems.asp">
<input type="hidden" name="page" value="">
<% 
varx = ""
for each Item in Request.Form 
	varx = varx & "&" & Item & "=" & Request.Form(Item)
	If Item <> "page" Then %>
<input type="hidden" name="<%=Item%>" value="<%=Request.Form(Item)%>">
<%		End If
next 
for each Item in Request.QueryString 
	varx = varx & "&" & Item & "=" & Request.QueryString (Item)
	If Item <> "page" Then %>
<input type="hidden" name="<%=Item%>" value="<%=Request.QueryString (Item)%>">
<%		End If
next %>
</form>
<script>
var varx = '<%=varx%>'
</script>
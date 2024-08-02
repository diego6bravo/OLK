<% addLngPathStr = "ventas/" %>
<!--#include file="lang/search_Clients.asp" -->

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style type="text/css">
.style1 {
	text-align: center;
}
</style>
</head>

<%

set rs = Server.CreateObject("ADODB.recordset")
sql = "select DispPosDeb from OADM"
set rs = conn.execute(sql)
If rs("DispPosDeb") = "Y" Then DispPosDeb = 1 Else DispPosDeb = -1

ShowClientBal = myAut.HasAuthorization(94)
ShowSuppBal = myAut.HasAuthorization(95)


set rx = Server.CreateObject("ADODB.RecordSet")
set rxVal = Server.CreateObject("ADODB.RecordSet")
sql = 	"select T0.rowIndex, IsNull(T1.AlterRowName, T0.rowName) rowName, T0.rowField, T0.RowType, T0.RowTypeRnd, T0.RowTypeDec, T0.rowOP, T0.RowAlign " & _
		"from olkcardrep T0 " & _
		"left outer join OLKCardRepAlterNames T1 on T1.rowIndex = T0.rowIndex and T1.LanID = " & Session("LanID") & " " & _
		"where T0.rowAccess in ('T','V') and T0.rowOP in ('T','O') and T0.ShowAt in ('A', 'S') " & _
		"order by T0.rowOrder asc"
rx.open sql, conn, 3, 1   


If Request("adSearch") = "Y" Then
	SearchID = Request("ID")

	set rdSearch = Server.CreateObject("ADODB.RecordSet")

	sql = "select IgnoreGeneralFilter, Query from OLKCustomSearch where ObjectCode = 2 and ID = " & SearchID
	set rdSearch = conn.execute(sql)
	IgnoreGeneralFilter = rdSearch("IgnoreGeneralFilter") = "Y"
	SearchQuery = rdSearch("Query")
	rdSearch.close
	
	sql = "select VarID, Variable, [Type], DataType, MaxChar, NotNull " & _
			"from OLKCustomSearchVars where ObjectCode = 2 and ID = " & SearchID & " and [Type] <> 'S'"
	rdSearch.open sql, conn, 3, 1
End If

If Request("D3") <> "" Then order1 = Request("D3") else order1 = "cardcode"
If Request("D5") <> "" Then order2 = Request("D5") else order2 = "asc"

sqlx = "select T0.CardCode "
 
sqlx = sqlx & _
	   "from OCRD T0 " & _
	   "left outer join OCRY T1 on T1.Code = T0.Country " & _
	   "inner join OCRG T2 on T2.GroupCode = T0.GroupCode " & _
	   "where Deleted = 'N' " & getSearchClientsFilter & " order by " & order1 & " " & order2

rs.close
set rs = Server.CreateObject("ADODB.RecordSet")
set rp = Server.CreateObject("ADODB.RecordSet")

rp.open sqlx, conn, 3, 1
If rp.recordcount = 1 Then
	Session("UserName") = rp("CardCode")
	Response.Redirect "activeClient.asp"
Else
rp.PageSize = 40
nPageCount = rp.PageCount
If Request("Page") <> "" Then nPage = CLng(Request("Page")) Else nPage = 1
If Not rp.Eof then 
	rp.AbsolutePage = nPage
	CardCode = ""
	do while not (rp.eof Or rp.AbsolutePage <> nPage )
		If CardCode <> "" Then CardCode = CardCode & ", "
		CardCode = CardCode & "N'" & Replace(rp("CardCode"), "'", "''") & "'"
	rp.MoveNext
	loop
	
	sqlx = "declare @MainCur nvarchar(3) set @MainCur = (select top 1 MainCurncy from oadm) " & _
			" declare @SlpCode int set @SlpCode = " & Session("vendid") & _
			" declare @dbName nvarchar(100) set @dbName = N'" & Session("OlkDB") & "' " & _
			" declare @LanID int set @LanID = " & Session("LanID") & " " & _
			" select T0.CardCode, T0.CardType, "
		   
	If Not myAut.HasAuthorization(174) Then
		sqlx = sqlx & "Case When T0.SlpCode = " & Session("vendid") & " Then Case When T0.Currency not in ('##',@MainCur) Then BalanceFC Else Balance End Else 0 End Balance, " & _
					"Case When T0.SlpCode = " & Session("vendid") & " Then 'Y' Else 'N' End ShowBalance, "
	Else
		sqlx = sqlx & "Case When T0.Currency not in ('##',@MainCur) Then BalanceFC Else Balance End Balance, "
		If Not IsNull(myApp.AgentClientsFilter) Then
			sqlx = sqlx & " Case When T0.CardCode in (" & Replace(Replace(myApp.AgentClientsFilter, "@SlpCode", Session("vendid")), "@Type", 2) & ") Then 'N' Else 'Y' End "
		Else
			sqlx = sqlx & " 'Y' "
		End If
	End If
		   
	sqlx = sqlx & " ShowBalance ,Case When T0.Currency <> '##' Then T0.Currency Else @MainCur End Currency, T0.Phone1, T0.E_Mail, " & _
		   "OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRD', 'CardName', T0.CardCode, T0.CardName) CardName, " & _
			"(select OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCPR', 'Name', (select CntctCode from OCPR where CardCode = T0.CardCode and Name = T0.CntctPrsn), CntctPrsn)) CntctPrsn, " & _
			"OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRG', 'GroupName', T0.GroupCode, GroupName) GroupName, " & _
			"OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRY', 'Name', T0.Country, Name) CountryName, T0.Country "
			
			

	do  while not rx.eof
		sqlx = sqlx & ", "
		If rx("rowTypeRnd") = "Y" Then rowTypeRnd = "Convert(Char(1),Convert(int,(10 * rand())))+ + " Else rowTypeRnd = ""
		If rx("rowType") = "L" or rx("rowType") = "M" or rx("rowType") = "H" Then
			Select Case rx("rowTypeDec")
				Case "S"
					myDec = myApp.SumDec
				Case "P"
					myDec = myApp.PriceDec
				Case "R"
					myDec = myApp.RateDec
				Case "Q"
					myDec = myApp.QtyDec
				Case "%"
					myDec = myApp.PercentDec
				Case "M"
					myDec = myApp.MeasureDec
			End Select
		End If
		If rx("rowType") = "L" Then
			sqlx = sqlx & " OLKCommon.dbo.DBOLKCode" & Session("ID") & "('L'," & rowTypeRnd & "Convert(nvarchar(20),(" & Rx("rowField") & ")), " & myDec & ")" & " As '" & Replace(Rx("rowName"), "'", "''") & "'"
		ElseIf rx("rowType") = "M" Then
			sqlx = sqlx & " OLKCommon.dbo.DBOLKCode" & Session("ID") & "('M'," & rowTypeRnd & "Convert(nvarchar(20),(" & Rx("rowField") & ")), " & myDec & ")" & " As '" & Replace(Rx("rowName"), "'", "''") & "'"
		ElseIf rx("rowType") = "H" Then
			sqlx = sqlx & " OLKCommon.dbo.DBOLKCode" & Session("ID") & "('H'," & rowTypeRnd & "Convert(nvarchar(20),(" & Rx("rowField") & ")), " & myDec & ")" & " As '" & Replace(Rx("rowName"), "'", "''") & "'"
		ElseIf rx("rowType") = "F" Then
			sqlx = sqlx & Rx("rowField") & " As '" & Replace(Rx("rowName"), "'", "''") & "'"
		Else
			sqlx = sqlx & "(" & Rx("rowField") & ") As N'CustCol" & Replace(Rx("rowIndex"), "'", "''") & "'"
		End If
	rx.movenext
	loop
			
	sqlx = sqlx & "from OCRD T0 " & _
		   "left outer join OCRY T1 on T1.Code = T0.Country " & _
		   "inner join OCRG T2 on T2.GroupCode = T0.GroupCode " & _
		   "where CardCode in (" & CardCode & ") "
	sqlx = sqlx & " order by " & order1 & " " & order2
	sqlx = Replace(sqlx, "@CardCode", "T0.CardCode")
	sqlx = QueryFunctions(sqlx)
Else
	sqlx = "select 'A' where 1 = 2"
End If
rs.open sqlx, conn, 3, 1

  iNextCount = nPage
  iCurMax = nPageCount/15
  iCurNext = 0
  do while iNextCount > 0
  	iNextCount = iNextCount - 15
  	iCurNext = iCurNext + 1
  loop
  If iCurMax - CInt(iCurMax) > 0 Then iCurMax = CInt(iCurMax) + 1

  fromI = (iCurNext*15)-14
  toI = (iCurNext*15)

  If iCurMax <= iCurNext Then toI = nPageCount
  If nPage > nPageCount Then nPage = nPageCount
  If nPage < 1 Then nPage = 1

%>
<table border="0" cellpadding="0" width="100%" id="table1">
	<tr class="GeneralTlt">
		<td colspan="3"><%=getsearch_ClientsLngStr("LttlClientSearch")%></td>
	</tr>
	<% If nPageCount > 1 Then %>
	<tr class="SearchPage">
		<td colspan="3">
		<% doSearchClientsPages %>
		</td>
	</tr>
	<% End If %>
	<tr>
		<td colspan="3">
		<table border="0" cellpadding="0" width="100%" id="table2">
			<tr class="FirmTlt3">
				<td width="2%" align="center" style="cursor: hand;" onclick="javascript:doSort('CardType');" <% doSearch_ClientsSortBG("CardType")%>>&nbsp;<% doSearch_ClientsSortImg("CardType")%></td>
				<td align="center" colspan="2" width="11%" style="cursor: hand;" onclick="javascript:doSort('CardCode');" <% doSearch_ClientsSortBG("CardCode")%>><%=getsearch_ClientsLngStr("DtxtCode")%><% doSearch_ClientsSortImg("CardCode")%></td>
				<td align="center" style="cursor: hand;" onclick="javascript:doSort('CardName');" <% doSearch_ClientsSortBG("CardName")%>><%=getsearch_ClientsLngStr("DtxtName")%><% doSearch_ClientsSortImg("CardName")%></td>
				<td align="center" colspan="2" style="cursor: hand;" onclick="javascript:doSort('CntctPrsn');" <% doSearch_ClientsSortBG("CntctPrsn")%>><%=getsearch_ClientsLngStr("DtxtContact")%><% doSearch_ClientsSortImg("CntctPrsn")%></td>
				<td align="center" style="cursor: hand;" onclick="javascript:doSort('Phone1');" <% doSearch_ClientsSortBG("Phone1")%>><%=getsearch_ClientsLngStr("DtxtPhone")%><% doSearch_ClientsSortImg("Phone1")%></td>
				<td align="center" style="cursor: hand;" onclick="javascript:doSort('GroupName');" <% doSearch_ClientsSortBG("GroupName")%>><%=getsearch_ClientsLngStr("DtxtGroup")%><% doSearch_ClientsSortImg("GroupName")%></td>
				<td align="center" style="cursor: hand;" onclick="javascript:doSort('Name');" <% doSearch_ClientsSortBG("Name")%>><%=getsearch_ClientsLngStr("DtxtCountry")%><% doSearch_ClientsSortImg("Name")%></td>
				<% If rx.recordcount > 0 Then rx.movefirst
				do while not rx.eof %><td align="center"><%=rx("rowName")%></td><% rx.movenext
				loop %>
				<td align="center" colspan="2" style="cursor: hand;" onclick="javascript:doSort('Balance');" <% doSearch_ClientsSortBG("Balance")%>><%=getsearch_ClientsLngStr("DtxtBalance")%><% doSearch_ClientsSortImg("Balance")%></td>
			</tr>
			<% 
			If Not rs.Eof Then
			do while not rs.eof %>
			<tr class="GeneralTbl">
				<td width="2%"><b>
				<p align="center">
				<% Select Case rs("CardType")
					Case "C" %>
				<img src="ventas/images/icon_supplier.gif" alt="<%=txtClient%>">
				<%	Case "L" %>
				<img src="ventas/images/icon_lead.gif" alt="<%=getsearch_ClientsLngStr("DtxtLead")%>">
				<% Case "S" %>
				<img src="ventas/images/icon_client.gif" alt="<%=getsearch_ClientsLngStr("DtxtSupplier")%>">
				<% End Select%></b></td>
				<td width="2%">
				<p align="center">
				<a href="javascript:goOp('<%=Replace(myHTMLEncode(rs("CardCode")), "'", "\'")%>');">
				<img border="0" src="design/0/images/<%=Session("rtl")%>felcahSelect.gif" width="15" height="13" alt="<%=Replace(getsearch_ClientsLngStr("LtxtSelActClient"), "{0}", myHTMLEncode(rs("CardName")))%>"></a></td>
				<td><%=Server.HTMLEncode(rs("CardCode"))%>&nbsp;</td>
				<td><%=rs("CardName")%>&nbsp;</td>
				<td width="15" class="style1"><% If rs("E_Mail") <> "" Then %><a href="mailto:<%=rs("E_Mail")%>"><img border="0" src="images/mail.gif" alt="<%=getsearch_ClientsLngStr("LtxtMailTo")%>: <%=rs("E_Mail")%>"></a><% end if %></td>
				<td><%=rs("CntctPrsn")%>&nbsp;</td>
				<td><%=rs("Phone1")%>&nbsp;</td>
				<td><%=rs("GroupName")%>&nbsp;</td>
				<td align="center"><img src="images/country/pic.aspx?filename=<%=rs("Country")%>.gif&MaxHeight=15" alt="<%=rs("CountryName")%>">&nbsp;</td>
				<% If rx.recordcount > 0 Then rx.movefirst
				do while not rx.eof
				RowAlign = ""
				If Not IsNull(rx("RowAlign")) Then
					Select Case rx("RowAlign")
						Case "L"
							RowAlign = " align=""left"""
						Case "R"
							RowAlign = " align=""right"""
						Case "C"
							RowAlign = " align=""center"""
					End Select
				End If %><td<%=RowAlign%>><%=rs("CustCol" & rx("rowIndex"))%></td><% rx.movenext
				loop %>
				<td width="15" class="style1"><% If Request("excell") <> "Y" and rs("ShowBalance") = "Y" Then%><a href="javascript:goCXC('<%=Replace(myHTMLEncode(rs("CardCode")), "'", "\'")%>');"><img border="0" src="design/0/images/<%=Session("rtl")%>felcahSelect.gif" width="15" height="14"></a><% End If %></td>
				<td>
				<p align="right"><nobr>
				<% If (rs("CardType") = "S" and ShowSuppBal or (rs("CardType") = "C" or rs("CardType") = "L") and ShowClientBal) and rs("ShowBalance") = "Y" Then %>
				<% If CDbl(rs("balance"))*DispPosDeb >= 0 Then %><nobr><%=rs("currency")%>&nbsp;<%=FormatNumber(CDbl(rs("balance"))*DispPosDeb,myApp.SumDec)%></nobr>
				<% Else %><font color="red"><nobr>(<%=rs("currency")%>&nbsp;<%=FormatNumber(CDbl(rs("balance"))*-1*DispPosDeb,myApp.SumDec)%>)</nobr></font>
				<% End If %><% Else %>****<% End If %></nobr></td>
			</tr>
		  <% rs.movenext
		  loop
		  Else %>
		  <tr class="GeneralTbl">
		  	<td colspan="12" align="center"><%=getsearch_ClientsLngStr("DtxtNoData")%></td>
		  </tr>
		  <% End If %>
		</table>
		</td>
	</tr>
	<% If nPageCount > 1 Then %>
	<tr class="SearchPage">
		<td colspan="3">
		<% doSearchClientsPages %>
		</td>
	</tr>
	<% End If %>
</table>
<% Sub doSearchClientsPages %>
<table border="0" cellpadding="0" width="100%">
	<tr>
		<% If iCurNext > 1 Then %><td width="14" class="FirmTlt3">
		<a href="javascript:goPage(<%= ((iCurNext-1)*15) %>);">
		<img border="0" src="design/0/images/<%=Session("rtl")%>prevAll.gif" width="12" height="13" align="left"></a>
		</td><% End If %><% If nPage > 1 Then	%>
		<td width="14" class="FirmTlt3">
		<a href="javascript:goPage(<%= nPage - 1 %>);">
		<img border="0" src="design/0/images/<%=Session("rtl")%>prev.gif" width="12" height="13" align="left"></a></td><% End If %>
		<td class="FirmTlt3" dir="ltr">
		<p align="center">
		<% 
		If nPageCount > 1 then
			For I = fromI To toI
				If I = nPage Then %>
				<font size="3">
				<b><%= I %></b></font>
				<% Else %>
				<a class="LnkSearchPaginacion" href="javascript:goPage(<%= I %>);"><%= I %></a>
				<% End If
			Next
		end if %></td><% If nPage < nPageCount Then %>
		<td width="14" class="FirmTlt3"><a href="javascript:goPage(<%= nPage + 1 %>);">
		<img border="0" src="design/<%=SelDes%>/images/<%=Session("rtl")%>next.gif" width="12" height="13" align="right">
		</a>
		</td><% End If %><% If iCurNext < iCurMax Then %>
		<td width="14" class="FirmTlt3">
		  <a href="javascript:goPage(<%= (iCurNext*15)+1 %>);">
		  <img border="0" src="design/<%=SelDes%>/images/<%=Session("rtl")%>nextAll.gif" width="12" height="13" align="right"></a></td><% End If %>
	</tr>
</table>
<% End Sub %>
<form name="frmGPage" action="clientsSearch.asp" method="post">
<% For each itm in Request.Form
If itm <> "Page" and itm <> "submit" and itm <> "D3" and itm <> "D5" Then %>
<input type="hidden" name="<%=itm%>" value="<%=Request(itm)%>">
<% End If
Next %>
<% For each itm in Request.QueryString
If itm <> "Page" and itm <> "submit" and itm <> "D3" and itm <> "D5" Then %>
<input type="hidden" name="<%=itm%>" value="<%=Request(itm)%>">
<% End If
Next %>
<input type="hidden" name="Page" value="">
<input type="hidden" name="D3" value="<%=Request("D3")%>">
<input type="hidden" name="D5" value="<%=Request("D5")%>">
</form>
<script language="javascript">
function goPage(p) { document.frmGPage.Page.value = p; document.frmGPage.submit(); }
function doSort(c)
{
	document.frmGPage.D3.value = c;
	if ('<%=Request("D3")%>' == c)
	{
		if ('<%=Request("D5")%>' == 'asc')
			document.frmGPage.D5.value = 'desc';
		else
			document.frmGPage.D5.value = 'asc';
	}
	else
	{
		document.frmGPage.D5.value = 'asc';
	}
	document.frmGPage.Page.value = 1;
	document.frmGPage.submit();
}
function goCXC(CardCode)
{
	<% If myAut.HasAuthorization(24) Then %>
	doMyLink('ventas/gocxc.asp', 'c1=' + CardCode, '');
	<% Else %>
	alert('<%=getsearch_ClientsLngStr("DtxtNoAccessObj")%>'.replace('{0}', '<%=getsearch_ClientsLngStr("LtxtStateOfAcct")%>'));
	return;
	<% End If %>
}
</script>
<% 
End If
Sub doSearch_ClientsSortImg(c)
	If LCase(Request("D3")) = LCase(c) Then
		If Request("D5") = "asc" Then
			Response.Write "<img src=""images/arrow_up.gif"">"
		Else
			Response.Write "<img src=""images/arrow_down.gif"">"
		End If
	End If
End Sub 
Sub doSearch_ClientsSortBG(c)
	If LCase(Request("D3")) = LCase(c) Then Response.Write "class=""GeneralTblBold2HighLight"""
End Sub

Function getSearchClientsFilter 
	sqlFilter = ""
	If Request("CardType") <> "" Then
		sqlFilter = sqlFilter & " and CardType  = '" & Request("CardType") & "' "
	Else
		CardType = ""
		If myAut.HasAuthorization(23) Then CardType = "'C'"
		If myAut.HasAuthorization(74) Then 
			If CardType <> "" Then CardType = CardType & ", "
			CardType = CardType & "'S'"
		End If
		If myAut.HasAuthorization(75) Then 
			If CardType <> "" Then CardType = CardType & ", "
			CardType = CardType & "'L'"
		End If
		sqlFilter = sqlFilter & " and CardType in (" & CardType & ") "
	End If
	If Request("String") <> "" Then 
		sqlFilter = sqlFilter & " and (T0.CardCode like N'%" & saveHTMLDecode(Request("string"), False) & "%' or "
	
		If myApp.EnableCSearchByVatId Then
			sqlFilter = sqlFilter & " T0.VatIdUnCmp like N'%" & saveHTMLDecode(Request("string"), False) & "%' or "
		End If
	
		If myApp.EnableCSearchByLicTradNum Then
			sqlFilter = sqlFilter & " T0.LicTradNum like N'%" & saveHTMLDecode(Request("string"), False) & "%' or "
		End If
		sqlFilter = sqlFilter & "OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRD', 'CardName', T0.CardCode, T0.CardName) collate database_default like N'%" & saveHTMLDecode(Request("string"), False) & "%') "
	End If
	If not myAut.HasAuthorization(60) Then sqlFilter = sqlFilter & " and T0.SLPCode = " & Session("vendid") & " "
	If Request("CardCodeFrom") <> "" Then sqlFilter = sqlFilter & " and T0.CardCode >= N'" & saveHTMLDecode(Request("CardCodeFrom"), False) & "' "
	If Request("CardCodeTo") <> "" Then sqlFilter = sqlFilter & " and T0.CardCode <= N'" & saveHTMLDecode(Request("CardCodeTo"), False) & "' "

	If Request("GroupNameFrom") <> "" or Request("GroupNameTo") <> "" Then
		sqlFilter = sqlFilter & "and ( "
	
		If Request("GroupNameFrom") <> "" Then sqlFilter = sqlFilter & " T2.GroupName >= N'" & saveHTMLDecode(Request("GroupNameFrom"), False) & "' "
		If Request("GroupNameFrom") <> "" and Request("GroupNameTo") <> "" Then sqlFilter = sqlFilter & " and "
		If Request("GroupNameTo") <> "" Then sqlFilter = sqlFilter & " T2.GroupName <= N'" & saveHTMLDecode(Request("GroupNameTo"), False) & "' "
	
		sqlFilter = sqlFilter & 	" or T0.GroupCode in (select PK " & _
									"	from OMLT X0 " & _
									"	inner join MLT1 X1 on X1.TranEntry = X0.TranEntry " & _
									"	where TableName = 'OCRG' and FieldAlias = 'GroupName' and LangCode = OLKCommon.dbo.OLKGetSBOLang(" & Session("LanID") & ") "
	
		If Request("GroupNameFrom") <> "" Then sqlFilter = sqlFilter & " and Convert(nvarchar(100),Trans) >= N'" & saveHTMLDecode(Request("GroupNameFrom"), False) & "' "
		If Request("GroupNameTo") <> "" Then sqlFilter = sqlFilter & " and Convert(nvarchar(100),Trans) <= N'" & saveHTMLDecode(Request("GroupNameTo"), False) & "' "		
	
		sqlFilter = sqlFilter & ") ) "
	End If
	If Request("CountryFrom") <> "" or Request("CountryTo") <> "" Then
		sqlFilter = sqlFilter & " and ("
	
		If Request("CountryFrom") <> "" Then sqlFilter = sqlFilter & " T1.Name >= N'" & saveHTMLDecode(Request("CountryFrom"), False) & "' "
		If Request("CountryFrom") <> "" and Request("CountryTo") <> "" Then sqlFilter = sqlFilter & " and "
		If Request("CountryTo") <> "" Then sqlFilter = sqlFilter & " T1.Name <= N'" & saveHTMLDecode(Request("CountryTo"), False) & "' "		
	
		sqlFilter = sqlFilter & 	" or T1.Code in (select PK " & _
									"	from OMLT X0 " & _
									"	inner join MLT1 X1 on X1.TranEntry = X0.TranEntry " & _
									"	where TableName = 'OCRY' and FieldAlias = 'Name' and LangCode = OLKCommon.dbo.OLKGetSBOLang(" & Session("LanID") & ") "
					
		If Request("CountryFrom") <> "" Then sqlFilter = sqlFilter & " and Convert(nvarchar(100),Trans) >= N'" & saveHTMLDecode(Request("CountryFrom"), False) & "' "
		If Request("CountryTo") <> "" Then sqlFilter = sqlFilter & " and Convert(nvarchar(100),Trans) <= N'" & saveHTMLDecode(Request("CountryTo"), False) & "' "		
	
		sqlFilter = sqlFilter & ") ) "
	End If

	If Not IsNull(myApp.AgentClientsFilter) and not IgnoreGeneralFilter Then
		sqlFilter = sqlFilter & " and T0.CardCode not in (" & Replace(Replace(myApp.AgentClientsFilter, "@SlpCode", Session("vendid")), "@Type", 1) & ") "
	End If
	
	If Request("chkQryGroup") <> "" Then
		chkQryGroups = Split(Request("chkQryGroup"), ", ")
		Select Case Request("QryGroupOp")
			Case "A"
				qryOP = " AND "
			Case "O"
				qryOP = " OR "
		End Select

		strQryGroups = "("
		For i = 0 to UBound(chkQryGroups)
			If i > 0 Then strQryGroups = strQryGroups & qryOP
			strQryGroups = strQryGroups & "QryGroup" & chkQryGroups(i) & " = 'Y'"
		Next
		strQryGroups = strQryGroups & ") "
		
		Select Case Request("QryGroupOp2")
			Case "I"
				sqlFilter = sqlFilter & " and " & strQryGroups
			Case "N"
				sqlFilter = sqlFilter & " and not " & strQryGroups
		End Select
	End If
	
  sqlFilter = sqlFilter & " and  " & _
	"(T0.ValidFor = 'N' or T0.ValidFor = 'Y' and " & _
	"( " & _
	"	(T0.ValidFrom is null or DateDiff(day,T0.ValidFrom,getdate()) >= 0) " & _
	"	and  " & _
	"	(T0.ValidTo is null or DateDiff(day,getdate(),T0.ValidTo) >= 0) " & _
	")) " & _
	"and  " & _
	"(T0.FrozenFor = 'N' or T0.FrozenFor = 'Y' and " & _
	"( " & _
	"	(/*T0.FrozenFrom is null or*/ DateDiff(day,FrozenFrom,getdate()) < 0) " & _
	"	and  " & _
	"	(/*T0.FrozenTo is null or*/ DateDiff(day,getdate(),FrozenTo) < 0) " & _
	")) "

  'Finaliado filtro
  If Request("adSearch") = "Y" Then
  	'rdSearch.Filter = "Type = 'CL'"
  	do while not rdSearch.eof
  		strValue = Request("var" & rdSearch("VarID"))
  		
  		If strValue <> "" Then
	  		If (rdSearch("DataType") = "numeric" or rdSearch("DataType") = "int" or rdSearch("Type") = "CL") Then
				SearchQuery = Replace(SearchQuery, "@" & rdSearch("Variable"),strValue)
			ElseIf rdSearch("Datatype") = "datetime" Then
				SearchQuery = Replace(SearchQuery, "@" & rdSearch("Variable"), "Convert(datetime,'" & SaveSqlDate(strValue) & "',120)")
	  		Else
				SearchQuery = Replace(SearchQuery, "@" & rdSearch("Variable"), "N'" & saveHTMLDecode(strValue, False) & "'")
			End If
		Else
			SearchQuery = Replace(SearchQuery, "@" & rdSearch("Variable"), "NULL")
		End If
		
  	rdSearch.movenext
  	loop
  	If InStr(Trim(SearchQuery), "@SystemFilters") = 1 Then
		andPos = InStr(LCase(sqlFilter), "and")
		sqlFilter = Mid(sqlFilter, andPos+3, Len(sqlFilter)-andPos-3)
	  	SearchQuery = Replace(SearchQuery, "@SystemFilters", "and (" & sqlFilter & ")")
	Else
		If InStr(Trim(SearchQuery), "@SystemFilters") <> 0 Then
			andPos = InStr(LCase(sqlFilter), "and")
			sqlFilter = Mid(sqlFilter, andPos+3, Len(sqlFilter)-andPos-3)
	  		SearchQuery = "and " & Replace(SearchQuery, "@SystemFilters", "(" & sqlFilter & ")")
	  		Response.Write SearchQuery
		Else
			SearchQuery = "and " & SearchQuery
		End If
	End If
	
	SearchQuery = Replace(SearchQuery, "OCRD.", "T0.")
	SearchQuery = Replace(SearchQuery, "OCRY.", "T1.")
	SearchQuery = Replace(SearchQuery, "OCRG.", "T2.")
	
	SearchQuery = Replace(SearchQuery, "@SlpCode", Session("vendid"))
	SearchQuery = Replace(SearchQuery, "@branch", Session("branch"))
	
  	getSearchClientsFilter = SearchQuery
  Else
	  getSearchClientsFilter = sqlFilter
  End If
End Function %>
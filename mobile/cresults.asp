<!--#include file="lang/cresults.asp" -->
<%
set rx = Server.CreateObject("ADODB.recordset")
MainCur = myApp.MainCur
         

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
 		

CardType = ""
If myAut.HasAuthorization(23) Then CardType = "'C'"
If myAut.HasAuthorization(75) Then
	If CardType <> "" Then CardType = CardType & ", "
	CardType = CardType & "'L'"
End If

If Request("orden1") <> "" Then order1 = Request("orden1") else order1 = "CardCode"
If Request("orden2") <> "" Then order2 = Request("orden2") else order2 = "asc"

sqlx = "select UPPER(CardCode) As CardCode, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRD', 'CardName', T0.CardCode, T0.CardName) CardName, " & _
	  "Case Currency When '##' Then N'" & MainCur & "' Else Currency End Currency, Phone1, E_Mail, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCPR', 'Name', (select CntctCode from OCPR where CardCode = T0.CardCode and Name = T0.CntctPrsn), CntctPrsn) CntctPrsn, "
	  
If not myAut.HasAuthorization(174) Then
	sqlx = sqlx & " Case When SlpCode = " & Session("vendid") & " Then Balance Else 0 End Balance, Case When SlpCode = " & Session("vendid") & " Then 'Y' Else 'N' End  "
Else
	sqlx = sqlx & " Balance, "
	If Not IsNull(myApp.AgentClientsFilter) Then
		sqlx = sqlx & " Case When T0.CardCode in (" & Replace(Replace(myApp.AgentClientsFilter, "@SlpCode", Session("vendid")), "@Type", 2) & ") Then 'N' Else 'Y' End "
	Else
		sqlx = sqlx & " 'Y' "
	End If
End If 
	  
sqlx = sqlx & "ShowBalance from OCRD T0 " & _
	  "left outer join ocry T1 on T1.Code = T0.Country " & _
	  "inner join OCRG T2 on T2.groupcode = T0.groupcode " & _
	  "where CardType in (" & CardType & ") " & GetClientSearchFilter & _
	  " order by " & order1 & " " & order2
	  
rx.open sqlx, conn, 3, 1
rx.PageSize = 10
rx.CacheSize = 10
If Request("p") <> "" Then iCurPage = CInt(Request("p")) Else iCurPage = 1
iPageCount = rx.PageCount
%><div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
    <!-- fwtable fwsrc="Z:\topmanage\logos\originales\pocket_art.png" fwbase="pocket_artpieza1.gif" fwstyle="FrontPage" fwdocid = "742308039" fwnested=""0" -->
    <tr>
      <td bgcolor="#9BC4FF">
      <table border="0" cellpadding="0" bordercolor="#111111" width="100%" id="AutoNumber1">
        <tr>
          <td width="100%">
          <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><%=getcresultsLngStr("LtxtClientSearchResul")%> 
          </font></b></td>
        </tr>
        <tr>
          <td width="100%">
          <% 
          If Not rx.Eof Then
          rx.AbsolutePage = iCurPage
          For i = 1 to rx.PageSize %>
          <table border="0" cellpadding="0" cellspacing="1"  bordercolor="#111111" width="100%" id="AutoNumber2">
            <tr>
              <td align="left" bgcolor="#66A4FF" width="40%"><b>
              <font size="1" face="Verdana"><%=getcresultsLngStr("DtxtCode")%></font></b></td>
              <td align="left" bgcolor="#66A4FF" colspan="3"><b>
				<font face="Verdana" size="1"><%=getcresultsLngStr("DtxtClient")%></font></b></td>
            </tr>
            <tr>
              <td bgcolor="#82B4FF" width="40%"><font size="1" face="Verdana"><a href="javascript:goOp('<%=Replace(myHTMLEncode(Rx("CardCode")), "'", "\'")%>', 1);"><font color="#000000"><%=Rx("CardCode")%></font></a></td>
              <td bgcolor="#82B4FF" colspan="3"><font size="1" face="Verdana"><%=rx("CardName")%></font></td>
            </tr>
            <% If rx("CntctPrsn") <> "" or rx("Phone1") <> "" Then %>
            <tr>
              <td bgcolor="#82B4FF" width="40%"><p><font size="1" face="Verdana"><%=rx("CntctPrsn")%></font></td>
              <td bgcolor="#82B4FF" colspan="3"><font face="Verdana" size="1"><%=rx("Phone1")%></font></td>
            </tr>
            <% End If %>
            <tr>
              <td bgcolor="#82B4FF" width="40%"><p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><font face="Verdana" size="1"><% If rx("balance") <> "0" and rx("ShowBalance") = "Y" Then%><font size="1" color="<%
    			If rx("balance") < "0" Then
   				 Response.Write "red"
    			 Else
   				 Response.Write "black"
  				 End If%>"><nobr><%=rx("currency")%>&nbsp;<%=FormatNumber(rx("balance"),2)%></nobr></font><% end if %></font>
  				 </td>
                  <td bgcolor="#82B4FF" width="16%">
                  <p align="center">
                  <% if rx("E_Mail") <> "" then %>
                  <a href="mailto:<%=rx("E_Mail")%>"><img border="0" src="images/mail.gif" width="16" height="16"></a>
                  <% end if %>
                  </p>
                  </td>
                  <td bgcolor="#82B4FF" width="16%">
                  <p align="center">
                  <% If myApp.EnableORDR or myApp.EnableOQUT Then %>
                  <a target="_parent" href="javascript:goOp('<%=Replace(myHTMLEncode(Rx("CardCode")), "'", "\'")%>', 2);">
                  <img border="0" src="images/newdoc.gif"></a><% Else %>&nbsp;<% End If %>
                  </td>
                  <td bgcolor="#82B4FF" width="16%">
                  <p align="center">
                  <a href="javascript:goOp('<%=Replace(myHTMLEncode(Rx("CardCode")), "'", "\'")%>', 1);">
   				 <img border="0" src="images/ficha.gif"></a>
   				 </td>
   				 </tr>
   				</table>
   				 <% rx.movenext
	        If rx.eof then exit for
		    next  %>
          </td>
        </tr>
        <tr>
          <td width="100%" colspan="4">
			<table border="0" cellpadding="0" cellspacing="1" width="100%" id="table1">
				<tr>
					<td width="16">
					<% If iCurPage > 1 Then %><a href="javascript:goP(<%=iCurPage-1%>)"><img border="0" src="images/flecha_prev.gif" width="16" height="16"></a><% Else %>&nbsp;<% End If %></td>
					<td>
					<p align="center">
					<select name="cmbPage" size="1" onchange="javascript:goP(this.value)">
					<% For i = 1 to iPageCount %>
					<option value="<%=i%>" <% If i = iCurPage Then %>selected<% End If %>><%=i%></option>
					<% Next %>
					</select></td>
					<td width="16">
					<% If iCurPage < iPageCount Then %><a href="javascript:goP(<%=iCurPage+1%>)"><img border="0" src="images/flecha_next.gif" width="16" height="16"></a><% End If %></td>
				</tr>
			</table>
			</td>
        </tr>
        <tr>
          <td width="100%"></td>
        </tr>
        <% Else %>
        <tr>
          <td width="100%" align="center"><b><font size="1" face="Verdana"><%=getcresultsLngStr("DtxtNoData")%></font></b></td>
        </tr>
        <% End If %>
      </table>
      </td>
    </tr>
    </table>
  </center>
</div>
<script type="text/javascript">
function goP(p)
{
	document.frmPage.p.value = p;
	document.frmPage.submit();
}
function goOp(CardCode, Op)
{
	switch (Op)
	{
		case 1:
			document.frmOp.cmd.value='datos';
			document.frmOp.card.value = CardCode;
			break;
		case 2:
			document.frmOp.cmd.value = 'docgo';
			document.frmOp.c1.value = CardCode;
			break;
	}
	document.frmOp.submit();
}
</script>
<form name="frmPage" method="post" action="operaciones.asp">
<% For each itm in Request.QueryString %><input type="hidden" name="<%=itm%>" value="<%=Request.QueryString(itm)%>"><% Next %>
<% For each itm in Request.Form
If itm <> "p" Then %><input type="hidden" name="<%=itm%>" value="<%=Request.Form(itm)%>"><% 
End If
Next %><input type="hidden" name="p" value="">
</form>
<form name="frmOp" method="post" action="operaciones.asp">
<input type="hidden" name="cmd" value="">
<input type="hidden" name="card" value="">
<input type="hidden" name="c1" value="">
</form>
<%

Function GetClientSearchFilter
	sqlFilter = ""
	
	If Request("string") <> "" Then
		sqlFilter = sqlFilter & " and (CardCode like N'%" & saveHTMLDecode(Request("string"), False) & "%' or "
	  
		If myApp.EnableCSearchByVatId Then
			sqlFilter = sqlFilter & " T0.VatIdUnCmp like N'%" & saveHTMLDecode(Request("string"), False) & "%' or "
		End If
	
		If myApp.EnableCSearchByLicTradNum Then
			sqlFilter = sqlFilter & " T0.LicTradNum like N'%" & saveHTMLDecode(Request("string"), False) & "%' or "
		End If

		sqlFilter = sqlFilter & "OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRD', 'CardName', T0.CardCode, T0.CardName) collate database_default like N'%" & saveHTMLDecode(Request("string"), False) & "%') " 
	
	End If
	
	If Request("CardType") <> "" Then sqlFilter = sqlFilter & " and T0.CardType = '" & Request("CardType") & "' "
	
	If Request("CardCodeFrom") <> "" Then sqlFilter = sqlFilter & " and T0.CardCode >= N'" & saveHTMLDecode(Request("CardCodeFrom"), False) & "' "
	If Request("CardCodeTo") <> "" Then sqlFilter = sqlFilter & " and T0.CardCode <= N'" & saveHTMLDecode(Request("CardCodeTo"), False) & "' "
	
	If Request("GroupCode") <> "" Then sqlFilter = sqlFilter & " and T0.GroupCode = " & Request("GroupCode")
	If Request("Country") <> "" Then sqlFilter = sqlFilter & " and T0.Country = N'" & Request("Country") & "'"
	
	If Session("useraccess") = "U" Then
		If not myAut.HasAuthorization(60) Then sqlFilter = sqlFilter & " and SlpCode = " & Session("Vendid")
	End If
	

	If Request("GroupNameFrom") <> "" or Request("GroupNameTo") <> "" Then
		sqlFilter = sqlFilter & " and ( "
	
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
	  	SearchQuery = Replace(SearchQuery, "@SystemFilters", sqlFilter)
	Else
		If InStr(Trim(SearchQuery), "@SystemFilters") <> 0 Then
			sysFilPos = InStr(SearchQuery, "@SystemFilters")
			andFilPos = InStrRev(Mid(SearchQuery, 1, sysFilPos-1), "and")
			SearchQuery = Left(SearchQuery, andFilPos-1) & Mid(SearchQuery, sysFilPos, Len(SearchQuery)-sysFilPos+1)
	  		SearchQuery = "and " & Replace(SearchQuery, "@SystemFilters", sqlFilter)
		Else
			SearchQuery = "and " & SearchQuery
		End If
	End If

	SearchQuery = Replace(SearchQuery, "OCRD.", "T0.")
	SearchQuery = Replace(SearchQuery, "OCRY.", "T1.")
	SearchQuery = Replace(SearchQuery, "OCRG.", "T2.")
	
	SearchQuery = Replace(SearchQuery, "@SlpCode", Session("vendid"))
	SearchQuery = Replace(SearchQuery, "@branch", Session("branch"))
	
  	GetClientSearchFilter = SearchQuery
  Else
	  GetClientSearchFilter = sqlFilter
  End If	
	
End Function
%>
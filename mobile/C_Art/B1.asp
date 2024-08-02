<% addLngPathStr = "C_Art/" %>
<!--#include file="lang/B1.asp" -->
<%


Dim iPageSize       'How big our pages are
Dim iPageCount      'The number of pages we get back
Dim iPageCurrent    'The page we want to show
Dim strOrderBy      'A fake parameter used to illustrate passing them
Dim strSQL          'SQL command to execute
Dim objPagingConn   'The ADODB connection object
Dim objPagingRS     'The ADODB recordset object
Dim iRecordsShown   'Loop controller for displaying just iPageSize records
Dim I               'Standard looping var
Dim pagex1
Dim pagex2

iPageSize = 4

set rd = Server.CreateObject("ADODB.recordset")
set rver = Server.CreateObject("ADODB.recordset")

set rQryGroups = Server.CreateObject("ADODB.RecordSet")
sql = "select T0.ItmsTypCod, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITG', 'ItmsGrpNam', T0.ItmsTypCod, T1.ItmsGrpNam) Name " & _  
"from OLKSearchQryGroups T0 " & _  
"inner join OITG T1 on T1.ItmsTypCod = T0.ItmsTypCod " 
rQryGroups.open sql, conn, 3, 1

If Session("RetVal") <> "" Then 
	sql = "select Object from R3_ObsCommon..TLOG where LogNum = " & Session("RetVal")
	set rd = conn.execute(sql)
	objectID = CInt(rd(0))
Else
	objectID = -1
End If

sql = "Select Case When T1.CatalogFilterAgent = 'Y' Then T1.CatalogFilter Else null End CatalogFilter " & _
"From OLKCommon T0 " & _
"left outer join OLKClientsAccess T1 on T1.CardCode = N'" & Session("UserName") & "' "
SET RD = conn.execute(SQL)

DisItemQryGroup = myApp.CarArt
EnableMinInv = myApp.GetEnableMinInv and objectID <> 23
CatalogFilter = rd("CatalogFilter")

If Request("string") <> "" and Request("rdSearchAs") = "S" Then
	arrSearchStr = Split(saveHTMLDecode(Request("string"), False), " ")
	rateStr = ""
	For i = 0 to UBound(arrSearchStr)
		If arrSearchStr(i) <> "" Then
			If rateStr <> "" Then rateStr = rateStr & " + "
			rateStr = rateStr & " OLKCommon.dbo.OLKRateString(N'" & saveHTMLDecode(arrSearchStr(i), False) & "', RateStr)"						
		End If
	Next
End If

If Request.QueryString("page") = "" Then

	If Request("adSearch") = "Y" Then
		SearchID = Request("ID")
	
		set rdSearch = Server.CreateObject("ADODB.RecordSet")
	
		sql = "select IgnoreGeneralFilter, Query from OLKCustomSearch where ObjectCode = 4 and ID = " & SearchID
		set rdSearch = conn.execute(sql)
		IgnoreGeneralFilter = rdSearch("IgnoreGeneralFilter") = "Y"
		SearchQuery = rdSearch("Query")
		rdSearch.close
		
		sql = "select VarID, Variable, [Type], DataType, MaxChar, NotNull " & _
				"from OLKCustomSearchVars where ObjectCode = 4 and ID = " & SearchID & " and [Type] <> 'S'"
		rdSearch.open sql, conn, 3, 1
	End If

	iPageCurrent = 1
	
	
	strPrice = "IsNull(Price, 0)"
	If Request("slist") <> "Y" Then strPrice = "IsNull(DisPrice, IsNull(Price, 0))"
	
	sqlstmt = ""
	
	If Request("string") <> "" and Request("rdSearchAs") = "S" Then
		sqlstmt = "select ItemCode, " & rateStr & " [Rate] from ("
	End If
	
	sqlstmt = sqlstmt & "select OITM.ItemCode "
	
	If Request("string") <> "" and Request("rdSearchAs") = "S" Then
		sqlstmt = sqlstmt & ", OITM.ItemCode + ' ' + OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'ItemName', OITM.ItemCode, ItemName) + IsNull(' ' + frgnName, '') + IsNull(' ' + Convert(nvarchar(4000),IsNull(F1.Trans, OMRC.FirmName)), '') + IsNull(' ' + Convert(nvarchar(4000),IsNull(G1.Trans, OITB.ItmsGrpNam)), '') "
		If myApp.EnableSearchAlterCode Then sqlstmt = sqlstmt & " + IsNull(' ' + Substitute, '') "
		If Not rQryGroups.Eof Then
			do while not rQryGroups.eof
				qryGroup = rQryGroups(0)
				sqlstmt = sqlstmt & Replace(" + Case When QryGroup{0} = 'Y' Then N' " & Replace(rQryGroups("Name"), "'", "''") & "' Else '' End ", "{0}", qryGroup)
			rQryGroups.movenext
			loop
			rQryGroups.movefirst
		End If
		
		sqlstmt = sqlstmt & " RateStr "
	End If
	
	sqlstmt = sqlstmt & " FROM OITM " & _
				"INNER JOIN ITM1 ON OITM.ItemCode = ITM1.ItemCode and ITM1.PriceList = " & Session("PList") & " " & _
		      	"left outer join OSCN on OSCN.ItemCode = OITM.ItemCode and OSCN.CardCode = N'" & saveHTMLDecode(Session("username"), False) & "' "
		      	

	  If myApp.GetMinInvBy = "W" or CustomInvVer Then
	  	sqlstmt = sqlstmt & "inner join OITW OITW on OITW.ItemCode = OITM.ItemCode and OITW.WhsCode = " & GetWhsCode("OITM") & " "
	  End If
		      	
		sqlstmt = sqlstmt & "left outer join OITB on OITB.ItmsGrpCod = OITM.ItmsGrpCod "
		If Request("string") <> "" and Request("rdSearchAs") = "S" or Request("ItmsGrpNamFrom") <> "" or Request("ItmsGrpNamTo") <> "" Then
			sqlstmt = sqlstmt & "left outer join OMLT G0 on G0.TableName = 'OITB' and G0.FieldAlias = 'ItmsGrpNam' and G0.PK = OITM.ItmsGrpCod " & _  
								"left outer join MLT1 G1 on G1.TranEntry = G0.TranEntry and G1.LangCode = OLKCommon.dbo.OLKGetSBOLang(" & Session("LanID") & ") "
		End If
		
		sqlstmt = sqlstmt & "left outer join OMRC on OMRC.FirmCode = OITM.FirmCode "
		If Request("string") <> "" and Request("rdSearchAs") = "S" or Request("FirmNameFrom") <> "" or Request("FirmNameTo") <> "" Then
			sqlstmt = sqlstmt & "left outer join OMLT F0 on F0.TableName = 'OMRC' and F0.FieldAlias = 'FirmName' and F0.PK = OITM.FirmCode " & _  
								"left outer join MLT1 F1 on F1.TranEntry = F0.TranEntry and F1.LangCode = OLKCommon.dbo.OLKGetSBOLang(" & Session("LanID") & ") "
		End If
		

		  If Request("slist") <> "Y" Then
		  sqlstmt = sqlstmt & _
			"left outer join (select X0.ItemCode,  " & _
			"Case X0.AutoUpdt When 'N' Then X0.DisPrice When 'Y' Then X0.ItmPrice-((X0.ItmPrice*X0.Discount)/100) End DisPrice,  " & _
			"X0.ItmPrice BefPrice, X0.DisPList " & _
			"from " & _
			"( " & _
			"  select P0.ItemCode, " & _
			"  Case P0.Expand When 'N' Then P0.AutoUpdt When 'Y' Then P1.AutoUpdt End AutoUpdt, " & _
			"  Case P0.Expand When 'N' Then P0.Discount When 'Y' Then P1.Discount End Discount, " & _
			"  Case P0.Expand When 'N' Then P0.Price When 'Y' Then P1.Price End DisPrice, " & _
			"  P2.Price ItmPrice, " & _
			"  Case P0.Expand When 'N' Then P0.Currency Else P1.Currency End Currency, Case P0.Expand When 'N' Then P0.ListNum Else P1.ListNum End DisPList " & _
			"  from OSPP P0 " & _
			"  left outer join SPP1 P1 on P1.ItemCode = P0.ItemCode and P1.CardCode = P0.CardCode " & _
			"  left outer join ITM1 P2 on P2.ItemCode = P0.ItemCode and P2.PriceList = Case P0.Expand When 'N' Then P0.ListNum Else P1.ListNum End " & _
			"  inner join OITM P3 on P3.ItemCode = P0.ItemCode " & _
			"  where P0.CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "' and " & _
			"  (P1.FromDate is null or DateDiff(day,getdate(),P1.FromDate) <= 0) and " & _
			"  (P1.ToDate is null or DateDiff(day,getdate(),P1.ToDate) >= 0) " & _
			"  union " & _
			"  select P0.ItemCode, 'Y' AutoUpdt, P1.Discount, null DisPrice, P2.Price ItmPrice, P2.Currency, " & Session("PList") & " DisPList " & _
			"  from OITM P0 " & _
			"  inner join OSPG P1 on P1.CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "' and " & _
			"  ( " & _
			"    (P1.ObjType = 52 and P0.ItmsGrpCod = P1.ObjKey) or " & _
			"    (P1.ObjType = 43 and P0.FirmCode = P1.ObjKey) " & _
			"  ) " & _
			"  inner join ITM1 P2 on P2.ItemCode = P0.ItemCode and P2.PriceList = " & Session("PList") & " " & _
			"  where P2.Price <> 0 and not exists " & _
			"  ( " & _
			"    select 'A'  " & _
			"    from OSPP S0  " & _
			"    left outer join SPP1 S1 on S1.ItemCode = S0.ItemCode and S1.CardCode = S0.CardCode  " & _
			"    where S0.CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "' and S0.ItemCode = P0.ItemCode " & _
			"    and " & _
			"    (S1.FromDate is null or DateDiff(day,getdate(),S1.FromDate) <= 0) and " & _
			"    (S1.ToDate is null or DateDiff(day,getdate(),S1.ToDate) >= 0) " & _
			"  )  " & _
			"  union " & _
			"  select P1.ItemCode, " & _
			"    P1.AutoUpdt, P1.Discount, P1.Price DisPrice, P2.Price ItmPrice, " & _
			"    P1.Currency Currency, " & Session("PList") & " DisPList " & _
			"  from OITM P0 " & _
			"  inner join SPP1 P1 on P1.ItemCode = P0.ItemCode " & _
			"  inner join ITM1 P2 on P2.ItemCode = P0.ItemCode and P2.PriceList = P1.ListNum " & _
			"  where P1.CardCode = N'*" & Session("PList") & "' " & _
			"  and not exists " & _
			"  ( " & _
			"    select 'A'  " & _
			"    from OSPP S0  " & _
			"    left outer join SPP1 S1 on S1.ItemCode = S0.ItemCode and S1.CardCode = S0.CardCode  " & _
			"    where S0.CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "' and S0.ItemCode = P0.ItemCode " & _
			"    and " & _
			"    (S1.FromDate is null or DateDiff(day,getdate(),S1.FromDate) <= 0) and " & _
			"    (S1.ToDate is null or DateDiff(day,getdate(),S1.ToDate) >= 0) " & _
			"  ) " & _
			"  and P1.ItemCode not in (select ItemCode collate database_default from OLKCommon.dbo.DBOLKGetCrdDiscGrpItm" & Session("ID") & "(N'" & saveHTMLDecode(Session("UserName"), False) & "'))  and " & _
					"(P1.FromDate is null or DateDiff(day,getdate(),P1.FromDate) <= 0) and " & _
					"(P1.ToDate is null or DateDiff(day,getdate(),P1.ToDate) >= 0) " & _
			") X0) D0 on D0.ItemCode = OITM.ItemCode "
		  End If

	If myApp.EnableCodeBarsQry and myApp.CodeBarsQryMethod = "I" Then
		sqlstmt = sqlstmt & "cross join (" & Replace(myApp.CodeBarsQry, "@CodeBars", "N'" & saveHTMLDecode(Request("string"), False) & "'") & ") tCodeBars "
	End If

	sqlstmt = sqlstmt & "where OITM.SellItem = 'Y' and OITM.Canceled = 'N' and OITM.TreeType <> 'T' "
	
	If DisItemQryGroup <> -1 Then sqlstmt = sqlstmt & " and QryGroup" & DisItemQryGroup & " = 'N' "
	
	sqlFilter = ""

	If Not IsNull(CatalogFilter) and CatalogFilter <> "" Then
		CatalogFilter = Replace(CatalogFilter, "@CardCode", "N'" & saveHTMLDecode(Session("UserName"), False) & "'")
		sqlFilter = sqlFilter & " and OITM.ItemCode not in (" & CatalogFilter & ") "
	End If
  
	If myApp.GetApplyGenFilter and not IgnoreGeneralFilter Then
		sqlFilter = sqlFilter & " and OITM.ItemCode not in (" & myApp.GetGenFilter & ") "
	End If 
  
	If Request("grupo") <> "" Then
		sqlFilter = sqlFilter & " and OITM.ItmsGrpCod = " & Request("grupo") & " "
	End If
	
	If Request("marca") <> "" Then 
		sqlFilter = sqlFilter & " and OITM.FirmCode = " & Request("marca") & " "
	End If 
	
	If Request("ItemCodeFrom") <> "" Then sqlFilter = sqlFilter & " and OITM.ItemCode >= N'" & saveHTMLDecode(Request("ItemCodeFrom"), False) & "' "
	If Request("ItemCodeTo") <> "" Then sqlFilter = sqlFilter & " and OITM.ItemCode <= N'" & saveHTMLDecode(Request("ItemCodeTo"), False) & "' "
 
  If Request("ItmsGrpNamFrom") <> "" Then sqlFilter = sqlFilter & " and Convert(nvarchar(4000),IsNull(G1.Trans, OITB.ItmsGrpNam)) >= N'" & saveHTMLDecode(Request("ItmsGrpNamFrom"), False) & "' "
  If Request("ItmsGrpNamTo") <> "" Then sqlFilter = sqlFilter & " and Convert(nvarchar(4000),IsNull(G1.Trans, OITB.ItmsGrpNam)) <= N'" & saveHTMLDecode(Request("ItmsGrpNamTo"), False) & "' "
	
  If Request("FirmNameFrom") <> "" Then sqlFilter = sqlFilter & " and Convert(nvarchar(4000),IsNull(F1.Trans, OMRC.FirmName() >= N'" & saveHTMLDecode(Request("FirmNameFrom"), False) & "' "
  If Request("FirmNameTo") <> "" Then sqlFilter = sqlFilter & " and Convert(nvarchar(4000),IsNull(F1.Trans, OMRC.FirmName)) <= N'" & saveHTMLDecode(Request("FirmNameTo"), False) & "' "
  
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
 
	If Request("string") <> "" Then
		Select Case Request("rdSearchAs")
			Case "S" 
			  	sqlSearchFilter = ""
			  	sqlSearchDesign = " (OITM.ItemCode like N'%{0}%' or OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'ItemName', OITM.ItemCode, ItemName) collate database_default " & _
			  						"like N'%{0}%' or frgnName like N'%{0}%' or Convert(nvarchar(4000),IsNull(F1.Trans, OMRC.FirmName)) collate database_default like N'%{0}%' or Convert(nvarchar(4000),IsNull(G1.Trans, OITB.ItmsGrpNam)) collate database_default like N'%{0}%' "
			  						
				If myApp.EnableSearchAlterCode Then sqlSearchDesign = sqlSearchDesign & "or Substitute like N'%{0}%'"
				
				If Not rQryGroups.Eof Then
					do while not rQryGroups.eof
						qryGroup = rQryGroups(0)
						sqlSearchDesign = sqlSearchDesign & Replace(" or Case When QryGroup{1} = 'Y' Then N'" & Replace(rQryGroups("Name"), "'", "''") & "' Else '' End like N'%{0}%' ", "{1}", qryGroup)
					rQryGroups.movenext
					loop
					rQryGroups.movefirst
				End If
	 											
				sqlSearchDesign = sqlSearchDesign & ") "
			  						
			  	For i = 0 to UBound(arrSearchStr)
			  		If arrSearchStr(i) <> "" Then
				  		If sqlSearchFilter <> "" Then sqlSearchFilter = sqlSearchFilter & " or "
				  		sqlSearchFilter = sqlSearchFilter & Replace(sqlSearchDesign, "{0}", arrSearchStr(i))
				  	End If
			  	Next
			  	sqlFilter = sqlFilter & " and ((" & sqlSearchFilter & ") or "
			  	
	  			If Not myApp.EnableCodeBarsQry Then
			  		sqlFilter = sqlFilter & "OITM.CodeBars = N'" & saveHTMLDecode(Request("String"), False) & "'"
			  	Else
					Select Case myApp.CodeBarsQryMethod
						Case "R"
						  	sqlFilter = sqlFilter & "OITM.CodeBars = (" & Replace(myApp.CodeBarsQry, "@CodeBars", "N'" & saveHTMLDecode(Request("String"), False) & "'") & ") "
						Case "I"
							sqlFilter = sqlFilter & "OITM.CodeBars = tCodeBars.CodeBars "
					End Select
		  			
			  	End If
			  	
				If myApp.EnableSearchItmSupp Then
					sqlFilter = sqlFilter & " or OITM.SuppCatNum = N'" & saveHTMLDecode(Request("String"), False) & "'"
				End If

			  	
			  	sqlFilter = sqlFilter & ") "
			Case "E" 
				sqlFilter = sqlFilter & " and (OITM.ItemCode = N'" & saveHTMLDecode(Request("string"), False) & "' or "
				
	
	  			If Not myApp.EnableCodeBarsQry Then
					sqlFilter = sqlFilter & "CodeBars = N'" & saveHTMLDecode(Request("string"), False) & "' "
			  	Else
					Select Case myApp.CodeBarsQryMethod
						Case "R"
						  	sqlFilter = sqlFilter & "OITM.CodeBars = (" & Replace(myApp.CodeBarsQry, "@CodeBars", "N'" & saveHTMLDecode(Request("String"), False) & "'") & ") "
						Case "I"
							sqlFilter = sqlFilter & "OITM.CodeBars = tCodeBars.CodeBars "
					End Select

		  			
			  	End If
			  	
				If myApp.EnableSearchItmSupp Then
					sqlFilter = sqlFilter & " or SuppCatNum = N'" & saveHTMLDecode(Request("String"), False) & "'"
				End If

	
				
				sqlFilter = sqlFilter & "or Substitute like N'%" & saveHTMLDecode(Request("string"), False) & "') "
		End Select
	End If


  If Request("chkProm") = "Y" Then
  	sqlFilter = sqlFilter & " and DisPrice is not null and DisPrice <> IsNull(Price, 0) and DisPList <> 0 "
  End If
  
  If Request("chkWL") = "Y" Then 
  	sqlFilter = sqlFilter & " and exists(select 'A' from OLKWL where CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "' and ItemCode = OITM.ItemCode) "
  End If
	
  If Request("new") = "ON" Then 
  	sqlFilter = sqlFilter & " and datediff(day, OITM.createdate, getdate()) <= " & myApp.f_creacion & " "
  End If

  If myApp.GetMinInvBy = "S" Then MinInvTbl = "OITM" Else MinInvTbl = "OITW"
  invFilter = ""
  If Request("InvFrom") <> "" Then 
	invFilter = invFilter & " " & MinInvTbl & ".OnHand >= " & Request("InvFrom") & " "
  ElseIf EnableMinInv Then ' and objectID <> 23
	invFilter = invFilter & " " & MinInvTbl & ".OnHand >= " & myApp.GetMinInv & " "
  End If
  If Request("InvTo") <> "" Then 
  	If invFilter <> "" Then invFilter = invFilter & " and"
  	invFilter = invFilter & " OITM.OnHand <= " & Request("InvTo") & " "
  End If
  
  If invFilter <> "" Then sqlFilter = sqlFilter & "and (" & invFilter & " or InvntItem = 'N') "
  
  If Request("pic") = "ON" Then 
  	sqlFilter = sqlFilter & " and PicturName is not null and RTrim(PicturName) <> ''"
  End If

	sqlFilter = sqlFilter & "and  " & _
		"(OITM.ValidFor = 'N' or OITM.ValidFor = 'Y' and " & _
		"( " & _
		"	(OITM.ValidFrom is null or DateDiff(day,OITM.ValidFrom,getdate()) >= 0) " & _
		"	and  " & _
		"	(OITM.ValidTo is null or DateDiff(day,getdate(),OITM.ValidTo) >= 0) " & _
		")) " & _
		"and  " & _
		"(OITM.FrozenFor = 'N' or OITM.FrozenFor = 'Y' and " & _
		"( " & _
		"	(/*OITM.FrozenFrom is null or*/ DateDiff(day,FrozenFrom,getdate()) < 0) " & _
		"	and  " & _
		"	(/*OITM.FrozenTo is null or*/ DateDiff(day,getdate(),FrozenTo) < 0) " & _
		")) "

  If Request("slist") <> "Y" Then
	If Request("PriceFrom") <> "" Then 
		sqlFilter = sqlFilter & " and " & strPrice & " >= " & getNumeric(Request("PriceFrom")) & " "
	Else
		sqlFilter = sqlFilter & " and " & strPrice & " >= " & myApp.MinPrice & " "
	End If
	If Request("PriceTo") <> "" Then 
		sqlFilter = sqlFilter & " and " & strPrice & " <= " & getNumeric(Request("PriceTo")) & " "
	End If
  End If

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
  	SearchQuery = Replace(SearchQuery, "@SystemFilters", "(" & Right(sqlFilter, Len(sqlFilter)-InStr(sqlFilter, "and")-2) & ")")
  	
	SearchQuery = Replace(SearchQuery, "@SlpCode", Session("vendid"))
	SearchQuery = Replace(SearchQuery, "@branch", Session("branch"))
	SearchQuery = Replace(SearchQuery, "@LanID", Session("LanID"))
	If Session("UserName") <> "" Then 
		SearchQuery = Replace(SearchQuery, "@CardCode", "N'" & saveHTMLDecode(Session("UserName"), False) & "'")
	Else
		SearchQuery = Replace(SearchQuery, "@CardCode", "NULL")
	End If
  	
  	sqlstmt = sqlstmt & " and " & SearchQuery
  Else
	  sqlstmt = sqlstmt & sqlFilter
  End If

	If Request("string") <> "" and Request("rdSearchAs") = "S" Then
		sqlstmt = sqlstmt & ") OITM order by [Rate] desc "
	End If
	Session("sqlstmt") = sqlstmt
	set rs = Server.CreateObject("ADODB.RecordSet")

	rs.open sqlstmt, conn, adOpenStatic, adLockReadOnly, adCmdText
Else
	sqlstmt = Session("sqlstmt")
	set rs = Server.CreateObject("ADODB.RecordSet")
	rs.open sqlstmt, conn, adOpenStatic, adLockReadOnly, adCmdText
	iPageCurrent = CInt(Request.QueryString("page"))
End If

If rs.recordcount = 1 then 
	If Request("slist") = "Y" Then nextCmd = "itemdetails&view=Y" Else nextCmd = "addcart"
	response.redirect "operaciones.asp?cmd=" & nextCmd & "&Item=" & CleanItem(rs("ItemCode")) & "&PackPrice=" & NoEmbalaje & "&retSearch=Y"
End If

RS.PageSize = iPageSize
RS.CacheSize = iPageSize
iPageCount = RS.PageCount
If iPageCurrent > iPageCount Then iPageCurrent = iPageCount
If iPageCurrent < 1 Then iPageCurrent = 1
      %>
<table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
	<form name="frmAddItems" method="post" action="operaciones.asp" onsubmit="return valFrm();">
    <tr>
      <td bgcolor="#9BC4FF">
      <table border="0" cellpadding="0" bordercolor="#111111" width="100%">
            <!--#include file="CardNameAdd.asp" -->
        <tr>
          <td>
          <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><%=getB1LngStr("LtxtItemsSearch")%> 
          </font></b></td>
        </tr>
          <%
If iPageCount = 0 Then %>
	<tr><td align="center"><b><font face="Verdana" size="1"><%=getB1LngStr("DtxtNoData")%></font></b></td></tr>
<% Else
RS.AbsolutePage = iPageCurrent

set rItm = Server.CreateObject("ADODB.RecordSet")
ItemCode = ""
for intRecord=1 to rs.PageSize
	If ItemCode <> "" Then ItemCode = ItemCode & ", "
	ItemCode = ItemCode & "N'" & mySearchString(rs("ItemCode")) & "'"
rs.MoveNext
if rs.EOF then exit for
next 
sqlAddStr = ""
If Request("slist") <> "Y" Then
	sqlAddStr = ", Case DisPList When 0 Then null Else DisPrice End DisPrice, IsNULL(Case DisPList When 0 Then DisPrice Else Price End,0) Price"
Else
	sqlAddStr = ", IsNULL(Price,0) Price "
End If
sql = ""

If Request("string") <> "" and Request("rdSearchAs") = "S" Then
	sql = "select *, " & rateStr & " [Rate] from ("
	sqlAddStr = sqlAddStr & ", OITM.ItemCode + ' ' + OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'ItemName', OITM.ItemCode, ItemName) + IsNull(' ' + frgnName, '') + IsNull(' ' + Substitute, '') RateStr "
End If
sql = 		sql & "select OITM.ItemCode, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'ItemName', OITM.ItemCode, OITM.ItemName) ItemName, OnHand, " & _
			"Currency, SalPackMsr, SalUnitMsr, CONVERT(Decimal(20,2),(OnHand - IsCommited + OnOrder) / NumInSale) AS ItmDisponible, SalPackUn, " & _
			"NumInSale, PicturName, OITM.DocEntry, DATEDIFF([day], CreateDate, GETDATE()) AS dias" & sqlAddStr & " " & _
			"FROM OITM INNER JOIN ITM1 ON OITM.ItemCode = ITM1.ItemCode and ITM1.PriceList = " & Session("PList") & " " & _
		    "left outer join OSCN on OSCN.ItemCode = OITM.ItemCode and OSCN.CardCode = N'" & saveHTMLDecode(Session("username"), False) & "' "

			
If Request("slist") <> "Y" Then
	sql = sql & _
			"left outer join (select X0.ItemCode,  " & _
			"Case X0.AutoUpdt When 'N' Then X0.DisPrice When 'Y' Then X0.ItmPrice-((X0.ItmPrice*X0.Discount)/100) End DisPrice,  " & _
			"X0.ItmPrice BefPrice, X0.DisPList " & _
			"from " & _
			"( " & _
			"  select P0.ItemCode, " & _
			"  Case P0.Expand When 'N' Then P0.AutoUpdt When 'Y' Then P1.AutoUpdt End AutoUpdt, " & _
			"  Case P0.Expand When 'N' Then P0.Discount When 'Y' Then P1.Discount End Discount, " & _
			"  Case P0.Expand When 'N' Then P0.Price When 'Y' Then P1.Price End DisPrice, " & _
			"  P2.Price ItmPrice, " & _
			"  Case P0.Expand When 'N' Then P0.Currency Else P1.Currency End Currency, Case P0.Expand When 'N' Then P0.ListNum Else P1.ListNum End DisPList " & _
			"  from OSPP P0 " & _
			"  left outer join SPP1 P1 on P1.ItemCode = P0.ItemCode and P1.CardCode = P0.CardCode " & _
			"  left outer join ITM1 P2 on P2.ItemCode = P0.ItemCode and P2.PriceList = Case P0.Expand When 'N' Then P0.ListNum Else P1.ListNum End " & _
			"  inner join OITM P3 on P3.ItemCode = P0.ItemCode " & _
			"  where P0.CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "' and " & _
			"  (P1.FromDate is null or DateDiff(day,getdate(),P1.FromDate) <= 0) and " & _
			"  (P1.ToDate is null or DateDiff(day,getdate(),P1.ToDate) >= 0) " & _
			"  union " & _
			"  select P0.ItemCode, 'Y' AutoUpdt, P1.Discount, null DisPrice, P2.Price ItmPrice, P2.Currency, " & Session("PList") & " DisPList " & _
			"  from OITM P0 " & _
			"  inner join OSPG P1 on P1.CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "' and " & _
			"  ( " & _
			"    (P1.ObjType = 52 and P0.ItmsGrpCod = P1.ObjKey) or " & _
			"    (P1.ObjType = 43 and P0.FirmCode = P1.ObjKey) " & _
			"  ) " & _
			"  inner join ITM1 P2 on P2.ItemCode = P0.ItemCode and P2.PriceList = " & Session("PList") & " " & _
			"  where P2.Price <> 0 and not exists " & _
			"  ( " & _
			"    select 'A'  " & _
			"    from OSPP S0  " & _
			"    left outer join SPP1 S1 on S1.ItemCode = S0.ItemCode and S1.CardCode = S0.CardCode  " & _
			"    where S0.CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "' and S0.ItemCode = P0.ItemCode " & _
			"    and " & _
			"    (S1.FromDate is null or DateDiff(day,getdate(),S1.FromDate) <= 0) and " & _
			"    (S1.ToDate is null or DateDiff(day,getdate(),S1.ToDate) >= 0) " & _
			"  )  " & _
			"  union " & _
			"  select P1.ItemCode, " & _
			"    P1.AutoUpdt, P1.Discount, P1.Price DisPrice, P2.Price ItmPrice, " & _
			"    P1.Currency Currency, " & Session("PList") & " DisPList " & _
			"  from OITM P0 " & _
			"  inner join SPP1 P1 on P1.ItemCode = P0.ItemCode " & _
			"  inner join ITM1 P2 on P2.ItemCode = P0.ItemCode and P2.PriceList = P1.ListNum " & _
			"  where P1.CardCode = N'*" & Session("PList") & "' " & _
			"  and not exists " & _
			"  ( " & _
			"    select 'A'  " & _
			"    from OSPP S0  " & _
			"    left outer join SPP1 S1 on S1.ItemCode = S0.ItemCode and S1.CardCode = S0.CardCode  " & _
			"    where S0.CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "' and S0.ItemCode = P0.ItemCode " & _
			"    and " & _
			"    (S1.FromDate is null or DateDiff(day,getdate(),S1.FromDate) <= 0) and " & _
			"    (S1.ToDate is null or DateDiff(day,getdate(),S1.ToDate) >= 0) " & _
			"  ) " & _
			"  and P1.ItemCode not in (select ItemCode collate database_default from OLKCommon.dbo.DBOLKGetCrdDiscGrpItm" & Session("ID") & "(N'" & saveHTMLDecode(Session("UserName"), False) & "'))  and " & _
					"(P1.FromDate is null or DateDiff(day,getdate(),P1.FromDate) <= 0) and " & _
					"(P1.ToDate is null or DateDiff(day,getdate(),P1.ToDate) >= 0) " & _
			") X0) D0 on D0.ItemCode = OITM.ItemCode "
End If
			
sql = sql & "where OITM.ItemCode in (" & ItemCode & ")"

If Request("string") <> "" and Request("rdSearchAs") = "S" Then
	sql = sql & ") OITM order by [Rate] desc "
End If

set rItm = conn.execute(sql)

If Request("slist") <> "Y" Then If myApp.AgentSaleUnit = 3 Then defQty = 1 Else defQty = FormatNumber(1, myApp.QtyDec)

do while not rItm.eof
          Codigo = rItm("ItemCode")
          Descripcion = rItm("ItemName")
          Inventario = rItm("OnHand")
          Cur = rItm("Currency")
          UnVenta = rItm("SalPackMsr")
          UnEmbalaje = rItm("SalUnitMsr")
          Disponible = rItm("ItmDisponible")
          NoVenta = rItm("SalPackUn")
          NoEmbalaje = rItm("NumInSale")
          If rItm("PicturName") <> "" Then
		  	Pic = rItm("PicturName")
		  Else
		  	Pic = "n_a.gif"
		  End If   
          Precio = CDbl(rItm("Price"))
		  PackPrice = CDbl(NoEmbalaje) * CDbl(Precio)
		  
		  hasDis = False
		  If Request("slist") <> "Y" Then
		  	If Not IsNull(rItm("DisPrice")) Then 
		  		If CDbl(rItm("Price")) > CDbl(rItm("DisPrice")) Then 
			  		hasDis = CDbl(rItm("Price")) <> CDbl(rItm("DisPrice"))
			  	Else
	                Precio = CDbl(rItm("DisPrice"))
	                PackPrice = CDbl(NoEmbalaje) * CDbl(rItm("DisPrice"))
			  	End If
		  	Else 
		  		hasDis = False
		  	End If
		  End If
%>
        <tr>
          <td>
			<table border="0" cellpadding="0" cellspacing="1"  bordercolor="#111111" width="100%">
            <tr>
              <td valign="top" rowspan="2">
              <table border="0" cellpadding="0" cellspacing="1"  bordercolor="#111111" width="100%">
                <tr>
                  <td colspan="2" bgcolor="#5197FF">
                  <b><font size="1" face="Verdana">&nbsp;<%=getB1LngStr("DtxtCode")%> - <%=getB1LngStr("DtxtDescription")%></font></b></td>
                  <td align="center" bgcolor="#5197FF" style="width: 40px;"><% If rItm("dias") <= myApp.f_creacion then %><img border="0" src="images/anewone.gif"><% End If %></td>
                </tr>
                <tr>
                  <td colspan="2" valign="top">
                  <font size="1" face="Verdana"><a href="operaciones.asp?cmd=<% If Request("slist") = "Y" Then %>itemdetails<% Else %>addcart<% End If %>&item=<%=Replace(Replace(Replace(rItm("ItemCode"),"#","%23"),"&","%26"),"""","%22")%>&view=Y"><b><%=Codigo%></b></a></td>
                  <td valign="top" bgcolor="#5197FF" style="width: 40px;">
                  <p align="center"><b>
                  <font size="1" face="Verdana"><%=getB1LngStr("LtxtAVL")%></font></b></td>
                </tr>
                <tr>
                  <td colspan="2" valign="top"><font size="1" face="Verdana"><%=Descripcion%>&nbsp;</font></td>
                  <td valign="top" style="width: 40px;"><font size="1" face="Verdana"><p align="center"><nobr><%=Disponible%></nobr></p></font></td>
                </tr>
                <tr>
                  <td align="center" bgcolor="#66A4FF"><b>
                  <font size="1" face="Verdana"><%=getB1LngStr("LtxtPrice")%><%=UnEmbalaje%></font></b></td>
                  <td align="center" bgcolor="#66A4FF"><b>
                  <font face="Verdana" size="1"><%=getB1LngStr("LtxtUnitPrice")%></font></b></td>
                  <td align="center" bgcolor="#66A4FF" style="width: 40px;"><b>
                  <font size="1" face="Verdana"><%=getB1LngStr("LtxtInventory")%></font></b></td>
                </tr>
                <tr>
                  <td align="center" bgcolor="#9BC4FF" dir="ltr">
                  <font size="1" face="Verdana"><% If hasDis Then %><strike><% End If %><%=Cur%>&nbsp;<%=FormatNumber(PackPrice,myApp.PriceDec)%><% If hasDis Then %></strike><% End If %></font></td>
                  <td align="center" bgcolor="#9BC4FF" dir="ltr">
                  <font face="Verdana" size="1"><% If hasDis Then %><strike><% End If %><%=Cur%>&nbsp;<%=FormatNumber(Precio,myApp.PriceDec)%><% If hasDis Then %></strike><% End If %></font></td>
                  <td align="center" bgcolor="#9BC4FF" style="width: 40px;">
                  <font size="1" face="Verdana"><nobr><%=Inventario%></nobr></font></td>
                </tr>
                <% If hasDis Then
                DisPrice = CDbl(rItm("DisPrice"))
                PackDisPrice = CDbl(NoEmbalaje) * CDbl(DisPrice) %>
                <tr>
                  <td align="center" bgcolor="#9BC4FF" dir="ltr">
                  <font size="1" face="Verdana"><%=Cur%>&nbsp;<%=FormatNumber(PackDisPrice,myApp.PriceDec)%></font></td>
                  <td align="center" bgcolor="#9BC4FF" dir="ltr">
                  <font face="Verdana" size="1"><%=Cur%>&nbsp;<%=FormatNumber(DisPrice,myApp.PriceDec)%></font></td>
                  <td align="center" bgcolor="#9BC4FF">
                  <font size="1" face="Verdana">&nbsp;</font></td>
                </tr>
                <% End If %>
              <% If Request("slist") <> "Y" Then %>
                <tr>
                  <td align="center" bgcolor="#9BC4FF"><input type="number" min="0" step="<%=GetNumberStep(myApp.QtyDec)%>" name="chkItemQty<%=rItm("DocEntry")%>" value="<%=defQty%>" onfocus="this.select();" onmouseup="event.preventDefault()"  onchange="javascript:chkQty(this, '<%=defQty%>', <%=myApp.QtyDec%>, 8699999999999.000)"></td>
                  <td align="center" bgcolor="#9BC4FF"><input type="checkbox" name="chkItem" id="chkItem" value="<%=rItm("DocEntry")%>"><input type="hidden" name="chkItemCode<%=rItm("DocEntry")%>" value="<%=Server.HTMLEncode(rItm("ItemCode"))%>"></td>
                  <td align="center" bgcolor="#9BC4FF" style="width: 40px;"><a href="operaciones.asp?cmd=addcart&Item=<%=Server.URLEncode(rItm("ItemCode"))%>&PackPrice=<%=NoEmbalaje%>">
				  <img border="0" src="images/addcart.gif" width="20" height="20" align="bottom"></a></td>
                </tr>
                <% End If %>
              </table>
              </td>
              <td valign="top" height="50" style="width: 64px;">
              <% If myApp.ShowPocketImg Then %>
                <p align="center">
                <font face="Garamond">
                <a href="operaciones.asp?cmd=viewImage&amp;FileName=<%=Pic%>">
                <img height="60" src="pic.aspx?filename=<%=Pic%>&dbName=<%=Session("OLKDB")%>" border="1" width="60"></a></font><% Else %>&nbsp;<% End If %></td>
            </tr>
            <tr>
              <td colspan="2"><hr color="#5197FF" size="1"></td>
            </tr>
          </table>
          <% rItm.movenext
          loop 
            end if%></td>
            <tr>
              <td>
              <table cellpadding="0" cellspacing="0" width="100%">
              	<tr>
              		<td>
		              <a href="operaciones.asp?cmd=slistsearch&slist=<%=Request("slist")%>"><b>
		              <img border="0" src="images/search_icon.gif" align="middle"><font size="1" face="Verdana"><%=getB1LngStr("LtxtNewSearch")%></font></b></a></td>
		           <% If iPageCount > 0 and Request("slist") <> "Y" Then %><td align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
		           	<input type="submit" name="btnAddItems" value="<%=getB1LngStr("DtxtAdd")%>">
		           </td><% End If %>
              	</tr>
              </table></td>
            </tr>
            <tr>
              <td width="100%">
              <table border="0" cellpadding="0" cellspacing="1" bordercolor="#111111" width="100%" dir="ltr">
                <tr>
                  <td width="8%" valign="top">
                    <% If iPageCurrent > 1 Then %><a href="javascript:goP(<%=iPageCurrent-1%>);"><img border="0" src="images/flecha_prev.gif"></a><% End If %></td>
                  <td width="85%" align="center">
                  <% If iPageCount > 0 Then %>
                  <select name="pageSelection" size="1" onchange="javascript:goP(this.value);" style="font-family: Verdana; font-size: 10px; border: 1px solid #5197FF; background-color: #9BC4FF">
                  	<% For I = 1 to iPageCount %>
                  	<option <% If I = iPageCurrent Then %>selected<% End If %> value="<%=i%>"><%=i%></option>
                  	<% Next %>
                  </select>
                  <% End If %>
					</td>
                  <td width="7%" valign="top">
                  <% If iPageCurrent < iPageCount Then %><a href="javascript:goP(<%=iPageCurrent+1%>);"><img border="0" src="images/flecha_next.gif"></a><% End If %></td>
                </tr>
              </table>
              </td>
            </tr>
      </table>
      </td>
    </tr>
    <input type="hidden" name="cmd" value="cartAddMulti">
	</form>
</table>
<script language="javascript">
var slist = '<%=Request("slist")%>';
var txtChkAtLead1Item = '<%=getB1LngStr("LtxtChkAtLead1Item")%>';
var txtValNumMaxVal = '<%=getB1LngStr("DtxtValNumMaxVal")%>';
var GetFormatDec = '<%=GetFormatDec()%>';
</script>
<script language="javascript" src="C_Art/B1.js"></script>
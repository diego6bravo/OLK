<!--#include file="clearItem.asp" -->
<% addLngPathStr = "" %>
<% If Request("PrintCatalog") = "Y" Then %>
<!--#include file="lang.asp" -->
<% End If %>
<!--#include file="lang/searchCart.asp" -->
<% If Request("excell") <> "Y" Then 

popH = 0
Select Case userType 
	Case "C"
		If Session("username") <> "-Anon-" Then
			popH = 450
		Else
			popH = 340
		End If
	Case "V"
		If Session("username") <> "" Then
			popH = 510
		Else
			popH = 420
		End If
End Select

PriceList = ""

If Request("CPList") <> "" Then
	PriceList = Request("CPList")
ElseIf Request("Excell") = "Y" Then
	PriceList = 0
Else
	PriceList = Session("PriceList")
End If
 %>
<SCRIPT LANGUAGE="JavaScript">
var txtChkAtLead1Item = '<%=getsearchCartLngStr("LtxtChkAtLead1Item")%>';
var UserType = '<%=userType%>';
var popH = <%=popH%>;
var txtValNumVal = '<%=getsearchCartLngStr("DtxtValNumVal")%>';
var QtyDec = <%=myApp.QtyDec%>;
var searchCmd = '<%=searchCmd%>';
var isAnon = <%=JBool(Session("UserName") = "-Anon-")%>;
var txtStartSesion = '<%=getsearchCartLngStr("LtxtStartSesion")%>';
</SCRIPT>
<script language="javascript" src="searchCart.js"></script>
<!--#include file="lcidReturn.inc"-->
<% 
	UserType = userType
	olkimgpath = Session("olkimgpath")
Else 
	userType = "V"
	UserType = "V"
	Session("branch") = Request("branch")
	Session("vendid") = Request("vendid")
End If

strOrden1 = Request("orden1")
strOrden2 = Request("orden2")

If strOrden1 = "" Then
	Select Case myApp.GetDefCatOrdr
		Case "C"
			strOrden1 = "OITM.ItemCode"
		Case Else
			strOrden1 = "ItemName"
	End Select
End If

rdSearchAs = Request("rdSearchAs")
If rdSearchAs = "" Then rdSearchAs = "S"
%>
<center>
<% 
set rs = Server.CreateObject("ADODB.recordset")
set rd = Server.CreateObject("ADODB.recordset")
set rver = Server.CreateObject("ADODB.recordset")
set rx = Server.CreateObject("ADODB.recordset")
set rp = Server.CreateObject("ADODB.recordset")

CatType = Request("Document")

If CatType = "" and userType = "C" Then
	If Request.Cookies("catMethod") <> "" Then CatType = Request.Cookies("catMethod") Else CatType = myApp.GetDefView 
End If

If Request("CPList") = "X" and Request("sourceDoc") <> "" and Request("sourceDoc") <> "-4" Then
	cardCodeFilterValue = "T0.CardCode"
ElseIf Session("UserName") <> "" and Session("UserName") <> "-Anon-" Then
	cardCodeFilterValue = "N'" & saveHTMLDecode(Session("UserName"), False) & "'"
Else
	cardCodeFilterValue = "NULL"
End If

If Session("RetVal") <> "" and Session("UserName") <> "-Anon-" Then 
	sql = "select Object from R3_ObsCommon..TLOG where LogNum = " & Session("RetVal")
	set rs = conn.execute(sql)
	objectID = CInt(rs(0))
	rs.close
Else
	objectID = -1
End If

sql = "select AliasID from CUFD where TableID = 'OITM' and TypeID = 'A' and EditType = 'I'"
set rd = conn.execute(sql)
If Not rd.Eof Then
	itmCaseImg = ""
	do while not rd.eof
		If itmCaseImg <> "" Then itmCaseImg = itmCaseImg & " or "
		itmCaseImg = itmCaseImg & "U_" & rd("AliasID") & " is not null "
	rd.movenext
	loop
	itmCaseImg = " Case When " & itmCaseImg & " Then 'Y' Else 'N' End "
Else
	itmCaseImg = " 'N' "
End If

SQL = "Select " & sqlAddStr & "ImgMaxSize, " & _
"catRows, "

If Request("PrintCatalog") = "Y" Then sql = sql & " pdfCols " 

sql = sql & "catCols, OLKCommon.dbo.OLKGetSBOLang(" & Session("LanID") & ") SBOLangID "

If Request("navIndex") <> "" Then
	sql = sql & ", NavQry NavFilter, ApplyAnonCatFilter  "
End If

If Request("sourceDoc") <> "" and Request("sourceDoc") <> "-4" Then	sql = sql & ", D0.T1, D0.T2 "

If (searchCmd <> "searchCatalog" or searchCmd = "searchCatalog" and PriceList <> "") and (optProm or userType = "V" and PriceList <> "X") then
	sql = sql & ", IsNull(P0.EncodeType, 'T') EncodeType, IsNull(P0.EncodeTypeRnd, 'N') EncodeTypeRnd "
End If

sql = sql & "From OLKCommon T0 "

If Request("navIndex") <> "" Then
	sql = sql & " inner join OLKCatNav T1 on T1.NavIndex = " & Request("navIndex") & " "
End If

If Request("sourceDoc") <> "" and Request("sourceDoc") <> "-4" Then	sql = sql & "cross join OLKDocConf D0 "

sql = sql & " cross join OADM A0 "

If (searchCmd <> "searchCatalog" or searchCmd = "searchCatalog" and PriceList <> "") and (optProm or userType = "V" and PriceList <> "X") then
	sql = sql & " left outer join OLKPriceListEncode P0 on P0.PriceList = " & PriceList & " "
End If

sql = sql & " cross join OLKCatOpt T2 Where T2.UserType = '" & UserType & "' and T2.CatType = '" & CatType & "'"


If Request("sourceDoc") <> "" and Request("sourceDoc") <> "-4" Then sql = sql & " and D0.ObjectCode = " & Request("sourceDoc")

SET RD = Conn.Execute(SQL)


If (searchCmd <> "searchCatalog" or searchCmd = "searchCatalog" and PriceList <> "") and (optProm or userType = "V" and PriceList <> "X") then
	PriceEncodeType = rd("EncodeType")
	PriceEncodeRnd = rd("EncodeTypeRnd")
End If

SBOLangID = rd("SBOLangID")
If myApp.SVer2005 and not myApp.SVer2007 Then TreePricOn = "'" & GetYN(myApp.TreePricOn) & "'" Else TreePricOn = "OITM.TreeType = 'S' and OITT.HideComp"
CustomInvVer = myApp.VerfyDisp = "C"
LogCat = myApp.EnableCSearchItemLog and userType = "C" and Session("UserName") <> "-Anon-"
LogSearch = myApp.EnableCSearchFilterLog and userType = "C" and Request("page") = "" and Session("UserName") <> "-Anon-"
If Request("navIndex") <> "" Then NavFilter = rd("NavFilter")
If Request("navIndex") <> "" Then ApplyAnonCatFilter = rd("ApplyAnonCatFilter") = "Y" Else ApplyAnonCatFilter = False

arty = "QryGroup" & myApp.CarArt
            
SaleType = myApp.GetSaleUnit

If Request("sourceDoc") <> "" and Request("sourceDoc") <> "-4" Then
	oTable = rd("T1")
	oTable1 = rd("T2")
	oTableS = Right(rd("T1"), 3)
End If

ImgMaxSize = rd("ImgMaxSize")


Select Case CatType
	Case "C"
		iPageSize = rd("catRows")*rd("catCols")
		catCols = rd("catCols")-1
	Case "T", "L"
		iPageSize = rd("catRows")
End Select

If Request("page") = "" Then
  If Request("cmd") = "searchCatalog" and Request("CPList") = "" then 
	Session("showPrice") = False 
  Else 
	Session("showPrice") = True
  End If
End If

If Request("string") <> "" and rdSearchAs = "S" Then
	arrSearchStr = Split(saveHTMLDecode(Request("string"), False), " ")
	rateStr = ""
	For i = 0 to UBound(arrSearchStr)
		If arrSearchStr(i) <> "" Then
			If rateStr <> "" Then rateStr = rateStr & " + "
			rateStr = rateStr & " OLKCommon.dbo.OLKRateString(N'" & saveHTMLDecode(arrSearchStr(i), False) & "', RateStr)"						
		End If
	Next
End If

set rQryGroups = Server.CreateObject("ADODB.RecordSet")
sql = "select T0.ItmsTypCod, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") &  ", 'OITG', 'ItmsGrpNam', T0.ItmsTypCod, T1.ItmsGrpNam) Name " & _  
"from OLKSearchQryGroups T0 " & _  
"inner join OITG T1 on T1.ItmsTypCod = T0.ItmsTypCod " 
rQryGroups.open sql, conn, 3, 1

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetObjectData" & Session("ID")
cmd.Parameters.Refresh()
cmd("@ObjType") = "S"
Select Case CatType
	Case "T"
		cmd("@ObjID") = 7
	Case "C"
		cmd("@ObjID") = 8
	Case "L"
		cmd("@ObjID") = 20
End Select
cmd("@UserType") = userType
set ra = cmd.execute()
strContent = ra("ObjContent")
strContent = Replace(strContent, "{SelDes}", SelDes)
If Session("rtl") <> "" Then
	strContent = Replace(strContent, "{rtl}", "rtl")
Else
	strContent = Replace(strContent, "{rtl}", "")
End If
strContent = Replace(strContent, "{rtl2}", Session("rtl"))
strContent = Replace(strContent, "{ImgMaxSize}", ImgMaxSize)
strContent = Replace(strContent, "{newItemAlt}", getsearchCartLngStr("LttlNewItm"))
If Session("rtl") = "" Then strContent = Replace(strContent, "{AlignRight}", "right") Else strContent = Replace(strContent, "{AlignRight}", "left")
If Request("dbName") = "" Then
	strContent = Replace(strContent, "{dbName}", Session("olkdb"))
Else
	strContent = Replace(strContent, "{dbName}", Request("dbName"))
End If
strContent = Replace(strContent, "{LtxtItmMoreImg}", getsearchCartLngStr("LtxtItmMoreImg"))

sql = 	"select T0.LineIndex, IsNull(T1.AlterColName, T0.ColName) ColName, T0.ColQuery, T0.ColType, T0.ColTypeRnd, T0.ColTypeDec, T0.ColAlign, T0.ColOrdr "

If CatType = "T" Then sql = sql & ", T0.ColIndex, (select count(ColIndex) from OLKTCart where ColOrdr = T0.ColOrdr and ColAccess in ('T','" & UserType & "')) ColCount "

sql = sql & 	"from OLK" & CatType & "Cart T0 " & _
				"left outer join OLK" & CatType & "CartAlterNames T1 on T1.LineIndex = T0.LineIndex and T1.LanID = " & Session("LanID") & " " & _
				"where T0.ColAccess in ('T','" & UserType & "') "

If (Request("CPList") = "" or Request("CPList") = "X") and Session("PriceList") = "" Then sql = sql& " and Convert(nvarchar(4000),colQuery) not like '%@PriceList%' "

If Session("RetVal") <> "" or ((Request("CPList") = "" or Request("CPList") = "X") and Request("DocNum") = "") Then sql = sql & " and Convert(nvarchar(4000),colQuery) not like '%@DocEntry%' and Convert(nvarchar(4000),colQuery) not like '%@LineNum%' and Convert(nvarchar(4000),colQuery) not like '%@Table%' "

If myApp.GetMinInvBy <> "W" and not CustomInvVer Then sql = sql & " and Convert(nvarchar(4000),colQuery) not like '%OITW.%' "

If Session("UserName") = "" and not (Request("CPList") = "X" and Request("sourceDoc") <> "" and Request("sourceDoc") <> "-4") Then
	sql = sql & " and Convert(nvarchar(4000), colQuery) not like N'%@CardCode%' "
ElseIf Session("UserName") = "-Anon-" Then
	sql = sql & " and ReqLogin = 'N' "
End If

sql = sql & "order by T0.ColOrdr"

If CatType = "T" Then sql = sql & ", T0.ColIndex"

rx.open sql, conn, 3, 1

If not rx.eof Then 

	do while not rx.eof

		If rx("ColType") = "L" or rx("ColType") = "M" or rx("ColType") = "H" Then

			Select Case rx("ColTypeDec")
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

		If rx("ColType") = "T" Then
			AddCode1 = ""
			AddCode2 = ""
		Else
			AddCode1 = "OLKCommon.dbo.DBOLKCode" & Session("ID") & "('" & rx("ColType") & "', "
			AddCode2 = ", " & myDec & ")"
		End If

		If rx("colTypeRnd") = "Y" Then 
			colTypeRnd1 = "Convert(Char(1),Convert(int,(10 * rand())))+ + Convert(nvarchar(20),("
			colTypeRnd2 = "))"
		Else
			colTypeRnd1 = ""
			colTypeRnd2 = ""
		End If
		
		colQuery = Replace(Replace(rx("ColQuery"),"@ItemCode", "OITM.ItemCode"), "@LanID", Session("LanID"))
		
		If Request("sourceDoc") <> "" and Request("sourceDoc") <> "-4" and Request("DocNum") <> "" Then
			colQuery = Replace(colQuery, "@Table", oTableS)
			colQuery = Replace(colQuery, "@DocNum", Request("DocNum"))
		End If

		AddFields = AddFields & AddCode1 & "(" & colTypeRnd1 & colQuery & colTypeRnd2 & ")" & AddCode2 & " As '" & Replace(rx("ColName"), "'", "''") & "', "
	rx.movenext
	loop
	rx.movefirst
End If

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

iSearchRecCount = 0

If Request("excell") <> "Y" and Request("PrintCatalog") <> "Y" Then 
  If Request("page") = "" Then iPageCurrent = 1 Else iPageCurrent = CInt(Request("page"))
  sqlstmt = GetSearchSqlStmt()
  rp.open sqlstmt, conn, 3, 1
  iSearchRecCount = rp.recordcount
End If
If Request("PrintCatalog") <> "Y" and Request("excell") <> "Y" Then
  rp.PageSize = iPageSize
  rp.CacheSize = iPageSize
  iPageCount = rp.PageCount

  iNextCount = iPageCurrent
  iCurMax = iPageCount/15
  iCurNext = 0
  do while iNextCount > 0
  	iNextCount = iNextCount - 15
  	iCurNext = iCurNext + 1
  loop
  If iCurMax - CInt(iCurMax) > 0 Then iCurMax = CInt(iCurMax) + 1

  fromI = (iCurNext*15)-14
  toI = (iCurNext*15)

  If iCurMax <= iCurNext Then toI = iPageCount
  If iPageCurrent > iPageCount Then iPageCurrent = iPageCount
  If iPageCurrent < 1 Then iPageCurrent = 1
Else
  searchCmd = Request("cmd")
  sqlFilter = getFilter
  If searchCmd <> "searchCatalog" or searchCmd = "searchCatalog" and Request("CPList") <> "" then 
  	If Not optProm and userType <> "V" or userType = "V" and Request("CPList") = "X" Then 
  		AddFields = AddFields & " 'N' WithoutPList, NULL "
  	Else
  		AddFields = AddFields & " Case DisPList When 0 Then 'Y' Else 'N' End WithoutPList, Case When IsNull(Price, 0) <> DisPrice Then DisPrice End "
  	End If
  	AddFields = AddFields & "DisPrice, "
  	If Not optProm and userType <> "V" or userType = "V" and Request("CPList") = "X" Then AddFields = AddFields & " NULL " Else AddFields = AddFields & "D0."
  	AddFields = AddFields & "ToDate, Case When " & TreePricOn & " = 'N' Then OLKCommon.dbo.DBOLKGetItemChildSum" & Session("ID") & "(@PriceList, " & cardCodeFilterValue & ", OITM.ItemCode) Else IsNull(Price, 0) End Price, "
  	
  	If (searchCmd <> "searchCatalog" or searchCmd = "searchCatalog" and Request("CPList") <> "") and (optProm or userType = "V" and Request("CPList") <> "X") then
  		AddFields = AddFields & " IsNull(IsNull(D0.Currency,T1.Currency), '') "
  	Else
	  	AddFields = AddFields & " IsNull(T1.Currency, '') "
  	End If
  	
  	AddFields = AddFields & " Currency, "
  	
  	If myApp.ShowPriceTax and userType = "C" and Session("RetVal") <> "" Then
		Select Case myApp.LawsSet 
			Case "PA", "AT", "AU", "BE", "CH", "CZ", "DE", "DK", "ES", "FI", "FR", "HU", "IT", "NL", "NO", "PL", "PT", "RU", "SE", "SK", "GB", "ZA"
				AddFields = AddFields & "		OLKCommon.dbo.DBOLKGetVatGroupRate" & Session("ID") &  "(Case PT6.VatStatus  " & _  
										"			When 'N'  " & _  
										"				Then Case PT6.CardType When 'S' Then 'C0' Else 'V0' End " & _  
										"			Else  " & _  
										"				Case When PT6.ECVatGroup is null Then " & _  
										"					Case PT6.CardType " & _  
										"						When 'S' Then OITM.VatGroupPu  " & _  
										"						Else OITM.VatGourpSa " & _  
										"					End " & _  
										"				Else PT6.ECVatGroup " & _  
										"				End  " & _  
								"		End, PT4.DocDate)/100 "
			Case "MX", "GT", "CR", "CL", "US", "CA"
				AddFields = AddFields & "		Case OITM.VatLiable When 'N' Then 0 Else (PT0.Rate/100) End "
			Case "IL"
				AddFields = AddFields & "		Case VatLiable When 'N' Then 0 When 'Y' Then " & myApp.VatPrcnt & "/100 End "
		End Select
		AddFields = AddFields & "	* ((100 - IsNull(PT3.Discount, 0)) / 100) AddVAT,  " 
  	End If
  End If
  
  If Request("PrintCatalog") <> "Y" and Request("excell") <> "Y" and Request("cmd") <> "searchCatalog" Then AddFields = AddFields & "Case When W0.ItemCode is null Then 'N' Else 'Y' End ChkWL, "
  
  If Request("sourceDoc") <> "" and Request("CPList") = "X" Then AddFields = AddFields & "T1.UseBaseUn, "
  sqlstmt = ""
  
  If Request("string") <> "" and rdSearchAs = "S" Then
  	sqlstmt = "select *, " & rateStr & " [Rate] from ("
  End If
  
  sqlstmt = sqlstmt & "select OITM.ItemCode, OITM.CreateDate, OITM.PicturName, " & _
  "Convert(nvarchar(max),IsNull(IsNull(F3.Trans, OITM.ItemName), '')) ItemName, " & _
  	"OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'SalPackMsr', OITM.ItemCode, OITM.SalPackMsr) SalPackMsr, " & _
  	"OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'SalUnitMsr', OITM.ItemCode, OITM.SalUnitMsr) SalUnitMsr, OITM.TreeType, " & _
  	"Replace(Convert(nvarchar(100),IsNull(OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'UserText', OITM.ItemCode, OITM.UserText) collate database_default, '')),Char(13),'<br>') + Case When Len(OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'UserText', OITM.ItemCode, OITM.UserText) collate database_default) > 100 Then '...' Else '' End UserText, " & _
  	AddFields & _
  "IsNull(OITM.NumInSale, 1) NumInSale, IsNull(OITM.SalPackUn, 1) SalPackUn, DateDiff(day, OITM.CreateDate,getdate()) Dias, " & _
  itmCaseImg & " MoreImg, " & _
  "Case When ManSerNum = 'Y' Then 'serial' "
  
  If myApp.Enable3dx Then
	  sqlstmt = sqlstmt & "	When IsNull((select U_GrpID from [@TM3DXITDE] where U_ItmEntry = OITM.DocEntry),-1) <> -1 and ManBtchNum = 'Y' Then '3dx' "
  End If
  
  sqlstmt = sqlstmt & "	When ManBtchNum = 'Y' Then 'batch' " & _
  "	Else 'item'  " & _
  "End ItemType, " & _
  "Case /*When ManSerNum = 'Y' Then N'" & saveHTMLDecode(getsearchCartLngStr("LtxtSerItem"), False) & "'*/ "
  
  If myApp.Enable3dx Then
	  sqlstmt = sqlstmt & "	When IsNull((select U_GrpID from [@TM3DXITDE] where U_ItmEntry = OITM.DocEntry),-1) <> -1 and ManBtchNum = 'Y' Then N'" & saveHTMLDecode(getsearchCartLngStr("Ltxt3dxItem"), False) & "' "
  End If
  
  
  sqlstmt = sqlstmt & "	When ManBtchNum = 'Y' Then N'" & saveHTMLDecode(getsearchCartLngStr("LtxtBtchItem"), False) & "' " & _
  "	Else N'" & getsearchCartLngStr("LtxtRegItem") & "'  " & _
  "End ItemTypeAlt "
  
  If Request("string") <> "" and rdSearchAs = "S" Then
  	sqlstmt = sqlstmt & ", OITM.ItemCode + ' ' + Convert(nvarchar(max),IsNull(IsNull(F3.Trans, OITM.ItemName), '')) + IsNull(' ' + frgnName, '') + IsNull(' ' + Convert(nvarchar(4000),IsNull(F1.Trans, OMRC.FirmName)), '') + IsNull(' ' + Convert(nvarchar(4000),IsNull(G1.Trans, OITB.ItmsGrpNam)), '') "
  	If myApp.EnableSearchAlterCode Then sqlstmt = sqlstmt & "+ IsNull(' ' + Substitute, '') "
	If Not rQryGroups.Eof Then
		do while not rQryGroups.eof
			qryGroup = rQryGroups(0)
			sqlstmt = sqlstmt & Replace(" + Case When OITM.QryGroup{0} = 'Y' Then N' " & Replace(rQryGroups("Name"), "'", "''") & "' Else '' End ", "{0}", qryGroup)
		rQryGroups.movenext
		loop
		rQryGroups.movefirst
	End If
  	sqlstmt = sqlstmt & " RateStr "
  End If
  
  sqlstmt = sqlstmt & getFrom & " where OITM.SellItem = 'Y' and OITM.Canceled = 'N' and OITM.TreeType <> 'T' " 
  
  If arty <> "QryGroup-1" Then
	  sqlstmt = sqlstmt & " and OITM." & arty & " = 'N' "
  End If
  
  sqlstmt = sqlstmt & sqlFilter
  
  If Request("string") <> "" and rdSearchAs = "S" Then
	  sqlstmt = sqlstmt & ") OITM order by [Rate] desc, " & strOrden1 & " " & strOrden2
  Else
	  sqlstmt = sqlstmt & " order by " & strOrden1 & " " & strOrden2
  End If
  
  If Request("CPList") <> "" Then
  	cpList = Request("CPList")
  	If cpList = "X" Then cpList = "-1"
	sqlstmt = Replace(sqlstmt,"@PriceList",cpList)
  ElseIf Request("Excell") = "Y" Then
	sqlstmt = Replace(sqlstmt,"@PriceList","0")
  Else
	sqlstmt = Replace(sqlstmt,"@PriceList",Session("PriceList"))
  End If
  sqlstmt = Replace(sqlstmt, "@SlpCode", Session("vendid"))
  sqlstmt = Replace(sqlstmt, "@CardCode", cardCodeFilterValue)
  rs.open QueryFunctions(sqlstmt), conn, 3, 1
  iPageCount = 1
  iSearchRecCount = rs.recordcount
End If
If Request("Excell") <> "Y" and Request("PrintCatalog") <> "Y" Then 
	If not rp.Eof Then
		rp.AbsolutePage = iPageCurrent
		ItemCode = ""
		for intRecord=1 to rp.PageSize
			If ItemCode <> "" Then ItemCode = ItemCode & ", "
			ItemCode = ItemCode & "N'" & Replace(rp("ItemCode"), "'", "''") & "'"
			If LogCat Then
				sqlLog = "EXEC OLKCommon..DBOLKAddItemLog" & Session("ID") & " N'" & saveHTMLDecode(Session("UserName"), False) & "', N'" & Replace(rp("ItemCode"), "'", "''") & "', 'S'"
	      		conn.execute(sqlLog)
			End If
		rp.MoveNext()
		If rp.EOF then exit for
		next
		  If searchCmd <> "searchCatalog" or searchCmd = "searchCatalog" and Request("CPList") <> "" then
		  	'If PriceEncodeType = "T" Then
			  	If Not optProm and userType <> "V" or userType = "V" and Request("CPList") = "X" Then 
			  		AddFields = AddFields & " Case When " & TreePricOn & " = 'N' Then OLKCommon.dbo.DBOLKGetItemChildSum" & Session("ID") & "(@PriceList, N'" & saveHTMLDecode(Session("UserName"), False) & "', OITM.ItemCode) Else IsNull(Price, 0) End Price, " & _
			  								"'N' WithoutPList, NULL "
			  	Else
			  		AddFields = AddFields & " Case When " & TreePricOn & " = 'N' Then OLKCommon.dbo.DBOLKGetItemChildSum" & Session("ID") & "(@PriceList, N'" & saveHTMLDecode(Session("UserName"), False) & "', OITM.ItemCode) Else IsNull(Case DisPList When 0 Then DisPrice Else Price End, 0) End  Price, " & _
			  		"Case DisPList When 0 Then 'Y' Else 'N' End WithoutPList, Case When IsNull(Price, 0) <> DisPrice and DisPList <> 0 Then DisPrice End "
			  	End If
			'Else
			'  	If Not optProm and userType <> "V" or userType = "V" and Request("CPList") = "X" Then 
			' 		AddFields = AddFields & " OLKCommon.dbo.DBOLKCode" & Session("ID") & "('" & PriceEncodeType & "',Case When " & TreePricOn & " = 'N' Then OLKCommon.dbo.DBOLKGetItemChildSum" & Session("ID") & "(@PriceList, N'" & saveHTMLDecode(Session("UserName"), False) & "', OITM.ItemCode) Else IsNull(Price, 0) End, " & myApp.PriceDec & ") Price, " & _
			'  								"'N' WithoutPList, NULL "
			'  	Else
			'  		AddFields = AddFields & " OLKCommon.dbo.DBOLKCode" & Session("ID") & "('" & PriceEncodeType & "',Case When " & TreePricOn & " = 'N' Then OLKCommon.dbo.DBOLKGetItemChildSum" & Session("ID") & "(@PriceList, N'" & saveHTMLDecode(Session("UserName"), False) & "', OITM.ItemCode) Else IsNull(Case DisPList When 0 Then DisPrice Else Price End, 0) End, " & myApp.PriceDec & ")  Price, " & _
			'  		"Case DisPList When 0 Then 'Y' Else 'N' End WithoutPList, Case When IsNull(Price, 0) <> DisPrice and DisPList <> 0 Then DisPrice End "
			'  	End If
			'End If
		  	AddFields = AddFields & "DisPrice, "
		  	If Not optProm and userType <> "V" or userType = "V" and Request("CPList") = "X" Then AddFields = AddFields & " NULL " Else AddFields = AddFields & "D0."
		  	AddFields = AddFields & " ToDate, "
		  	
		  	If (searchCmd <> "searchCatalog" or searchCmd = "searchCatalog" and Request("CPList") <> "") and (optProm or userType = "V" and Request("CPList") <> "X") then
		  		AddFields = AddFields & " IsNull(IsNull(D0.Currency,T1.Currency), '') "
		  	Else
			  	AddFields = AddFields & " IsNull(T1.Currency, '') "
		  	End If
		  	


			AddFields = AddFields & " Currency, "
		  	
		  	If myApp.ShowPriceTax and userType = "C" and Session("RetVal") <> "" Then
				Select Case myApp.LawsSet 
					Case "PA", "AT", "AU", "BE", "CH", "CZ", "DE", "DK", "ES", "FI", "FR", "HU", "IT", "NL", "NO", "PL", "PT", "RU", "SE", "SK", "GB", "ZA"
						AddFields = AddFields & "		OLKCommon.dbo.DBOLKGetVatGroupRate" & Session("ID") &  "(Case PT6.VatStatus  " & _  
												"			When 'N'  " & _  
												"				Then Case PT6.CardType When 'S' Then 'C0' Else 'V0' End " & _  
												"			Else  " & _  
												"				Case When PT6.ECVatGroup is null Then " & _  
												"					Case PT6.CardType " & _  
												"						When 'S' Then OITM.VatGroupPu  " & _  
												"						Else OITM.VatGourpSa " & _  
												"					End " & _  
												"				Else PT6.ECVatGroup " & _  
												"				End  " & _  
										"		End, PT4.DocDate)/100 "
					Case "MX", "GT", "CR", "CL", "US", "CA"
						AddFields = AddFields & "		Case OITM.VatLiable When 'N' Then 0 Else (PT0.Rate/100) End "
					Case "IL"
						AddFields = AddFields & "		Case VatLiable When 'N' Then 0 When 'Y' Then " & myApp.VatPrcnt & "/100 End "
				End Select
				AddFields = AddFields & "	* ((100 - IsNull(PT3.Discount, 0)) / 100) AddVAT,  " 
		  	End If
		  End If
  If Request("PrintCatalog") <> "Y" and Request("excell") <> "Y" and Request("cmd") <> "searchCatalog" Then AddFields = AddFields & "Case When W0.ItemCode is null Then 'N' Else 'Y' End ChkWL, "
		  If Request("sourceDoc") <> "" and Request("CPList") = "X" Then AddFields = AddFields & "T1.UseBaseUn, "
		  sqlstmt = ""
		  
		  If Request("string") <> "" and rdSearchAs = "S" Then
		  	sqlstmt = "select *, " & rateStr & " [Rate] from ("
		  End If
		  
		  sqlstmt = sqlstmt & "select OITM.ItemCode, OITM.CreateDate, OITM.PicturName, " & _
		  "Convert(nvarchar(max),IsNull(IsNull(F3.Trans, OITM.ItemName), '')) ItemName, " & _
			  "OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'SalPackMsr', OITM.ItemCode, OITM.SalPackMsr) SalPackMsr, " & _
			  "OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'SalUnitMsr', OITM.ItemCode, OITM.SalUnitMsr) SalUnitMsr, OITM.TreeType, " & _
			  "Replace(Convert(nvarchar(100),IsNull(OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'UserText', OITM.ItemCode, OITM.UserText) collate database_default, '')),Char(13),'<br>') + Case When Len(Convert(nvarchar(101),OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'UserText', OITM.ItemCode, OITM.UserText))) > 100 Then '...' Else '' End UserText, " & _
			  AddFields & _
		  "IsNull(OITM.NumInSale, 1) NumInSale, IsNull(OITM.SalPackUn, 1) SalPackUn, DateDiff(day, OITM.CreateDate,getdate()) Dias, " & _
		  "Case When ManSerNum = 'Y' Then 'serial' "
		  
		  If myApp.Enable3dx Then
			  sqlstmt = sqlstmt & "	When IsNull((select U_GrpID from [@TM3DXITDE] where U_ItmEntry = OITM.DocEntry),-1) <> -1 and ManBtchNum = 'Y' Then '3dx' "
		  End If
		  
		  sqlstmt = sqlstmt & "	When ManBtchNum = 'Y' Then 'batch' " & _
		  "	Else 'item'  " & _
		  "End ItemType, " & _
		  "Case /* When ManSerNum = 'Y' Then N'" & saveHTMLDecode(getsearchCartLngStr("LtxtSerItem"), False) & "'*/ "
		  
		  If myApp.Enable3dx Then
			  sqlstmt = sqlstmt & "	When IsNull((select U_GrpID from [@TM3DXITDE] where U_ItmEntry = OITM.DocEntry),-1) <> -1 and ManBtchNum = 'Y' Then N'" & saveHTMLDecode(getsearchCartLngStr("Ltxt3dxItem"), False) & "' "
		  End If
		  
		  sqlstmt = sqlstmt & "	When ManBtchNum = 'Y' Then N'" & saveHTMLDecode(getsearchCartLngStr("LtxtBtchItem"), False) & "' " & _
		  "	Else N'" & getsearchCartLngStr("LtxtRegItem") & "'  " & _
		  "End ItemTypeAlt, " & itmCaseImg & " MoreImg "
  
		  If Request("string") <> "" and rdSearchAs = "S" Then
		  	sqlstmt = sqlstmt & ", OITM.ItemCode + ' ' + Convert(nvarchar(max),IsNull(IsNull(F3.Trans, OITM.ItemName), '')) + IsNull(' ' + frgnName, '') + IsNull(' ' + Convert(nvarchar(4000),IsNull(F1.Trans, OMRC.FirmName)), '') + IsNull(' ' + Convert(nvarchar(4000),IsNull(G1.Trans, OITB.ItmsGrpNam)), '') "
		  	If myApp.EnableSearchAlterCode Then sqlstmt = sqlstmt & "+ IsNull(' ' + Substitute, '') "
			If Not rQryGroups.Eof Then
				do while not rQryGroups.eof
					qryGroup = rQryGroups(0)
					sqlstmt = sqlstmt & Replace(" + Case When OITM.QryGroup{0} = 'Y' Then N' " & Replace(rQryGroups("Name"), "'", "''") & "' Else '' End ", "{0}", qryGroup)
				rQryGroups.movenext
				loop
				rQryGroups.movefirst
			End If
		  	sqlstmt = sqlstmt & " RateStr "
		  End If
		  
		  sqlstmt = sqlstmt & getFrom & " where OITM.ItemCode in (" & ItemCode & ")"
		  
		  If Request("CPList") <> "" Then
		  	cpList = Request("CPList")
		  	If cpList = "X" Then cpList = "-1"
			sqlstmt = Replace(sqlstmt,"@PriceList",cpList)
		  ElseIf Request("Excell") = "Y" Then
			sqlstmt = Replace(sqlstmt,"@PriceList","0")
		  Else
			sqlstmt = Replace(sqlstmt,"@PriceList",Session("PriceList"))
		  End If
		  
		  If searchCmd = "searchCatalog" and Request("sourceDoc") <> "" then ' and Request("CPList") = "X"
			If Request("CPList") = "" or Request("CPList") = "X" Then
				If Request("sourceDoc") <> "-4" Then
					If Request("LinkRep") <> "Y" Then
						sqlstmt = sqlstmt & "and T0.DocNum = " & Request("DocNum") & " "
					Else
						sqlstmt = sqlstmt & "and T0.DocEntry = " & Request("DocNum") & " "
					End If
				Else
					sqlstmt = sqlstmt & "and T1.LogNum = " & Request("DocNum") & " "
				End If
			Else
				If Request("sourceDoc") <> "-4" Then
					sqlFilter = sqlFilter & " and OITM.ItemCode in " & _
					"(select itemcode from " & oTable1 & " X1 "
					
					If Request("LinkRep") <> "Y" Then
						sqlFilter = sqlFilter & "inner join " & oTable & " X0 on X0.DocEntry = X1.DocEntry where docnum = " & Request("DocNum")
					Else
						sqlFilter = sqlFilter & " where X1.DocEntry = " & Request("DocNum")
					End If
					
					sqlFilter = sqlFilter & ")"
				Else
					sqlFilter = sqlFilter & " and OITM.ItemCode in " & _
					"(select itemcode collate database_default from R3_ObsCommon..DOC1 X1 inner join " & _
					" R3_ObsCommon..TDOC X0 on X0.LogNum = X1.LogNum where X0.LogNum = " & Request("DocNum") & ")"
				End If
			End If
		  ElseIf searchCmd = "searchCatalog" and Request("CPList") = "" then
		  Else
		  End If
		  
		  	If Request("string") <> "" and rdSearchAs = "S" Then
			  	sqlstmt = sqlstmt & ") OITM order by [Rate] desc, " & strOrden1 & " " & strOrden2
		  	Else
			  	sqlstmt = sqlstmt & " order by " & strOrden1 & " " & strOrden2
			End If
		  
		  If Request("CPList") = "X" and Request("sourceDoc") <> "" and Request("sourceDoc") <> "-4" Then
			sqlstmt = Replace(sqlstmt,"@CardCode", "T0.CardCode")
		  ElseIf Session("UserName") <> "" Then
			sqlstmt = Replace(sqlstmt,"@CardCode", "N'" & saveHTMLDecode(Session("UserName"), False) & "'")
		  End If
		  
		  sqlstmt = Replace(sqlstmt, "@SlpCode", Session("vendid"))
		  
		  sqlstmt = QueryFunctions(sqlstmt)

		  rs.open sqlstmt, conn, 3, 1
	  Else
	  	sql = "select 'A' where 1 = 2"
		
		If InStr(1, sqlstmt, "order by") = 0 and strOrden1 <> "" Then sqlstmt = sqlstmt & " order by " & strOrden1 & " " & strOrden2

	  	rs.open sqlstmt, conn, 3, 1
	  End If
End If
%>

<%=doSearchTtl(Left(strContent, InStr(strContent, "<!--startPaging-->")-1))%>
<%=doPagingStr(getMid(strContent, "startPaging", "endPaging"))%>
<%=doChkItems%>
<% 
Select Case CatType
	Case "T"
		strStore = getMid(strContent, "startLoop", "endLoop")
		strStore = Replace(strStore, "{DtxtProm}", getsearchCartLngStr("DtxtProm"))
		strStore = Replace(strStore, "{LtxtExpires}", getsearchCartLngStr("LtxtExpires"))
		strStore = Replace(strStore, "{LtxtViewProd}", getsearchCartLngStr("LtxtViewProd"))
		strStore = Replace(strStore, "{LtxtWishList}", getsearchCartLngStr("LtxtWishList"))
		strStore = Replace(strStore, "{DtxtBuy}", getsearchCartLngStr("LtxtAdd2Cart"))
		strStore = Replace(strStore, "{DtxtTemplate}", getsearchCartLngStr("DtxtRecommendations"))
		If Not (myApp.ShowPriceTax and userType = "C" and Session("RetVal") <> "") Then
			strStore = Replace(strStore, "{DtxtPrice}", getsearchCartLngStr("DtxtPrice"))
		Else
			strStore = Replace(strStore, "{DtxtPrice}", getsearchCartLngStr("DtxtPrice") & "+" & txtTax)
		End If
		
		If Session("rtl") = "" Then
			strStore = Replace(strStore, "{rtl}", "")
			strStore = Replace(strStore, "{ItemCodeAlign}", "right")
		Else
			strStore = Replace(strStore, "{rtl}", "Rtl")
			strStore = Replace(strStore, "{ItemCodeAlign}", "left")
		End If
		If userType = "C" and InStr(strStore, "startItemTemplate") <> 0 Then
			strStore = Replace(strStore, getFullMid(strStore, "startItemTemplate", "endItemTemplate"), "")
		End If
		If InStr(strStore, "startItemOpt") <> 0 Then
			If Request("PrintCatalog") <> "Y" Then
				strStore = Replace(strStore, getFullMid(strStore, "startItemOpt", "endItemOpt"), getMid(strStore, "startItemOpt", "endItemOpt"))
			Else
				strStore = Replace(strStore, getFullMid(strStore, "startItemOpt", "endItemOpt"), "")
			End If
		End If
		If InStr(strStore, "startEditBtn") <> 0 Then
			If myAut.HasAuthorization(170) and searchCmd = "searchCatalog" and Request("PrintCatalog") <> "Y" Then
				strStore = Replace(strStore, getFullMid(strStore, "startEditBtn", "endEditBtn"), getMid(strStore, "startEditBtn", "endEditBtn"))
				strStore = Replace(strStore, "{DtxtEdit}", getsearchCartLngStr("DtxtEdit"))
				strStore = Replace(strStore, "{DtxtDuplicate}", getsearchCartLngStr("DtxtDuplicate"))
			Else
				strStore = Replace(strStore, getFullMid(strStore, "startEditBtn", "endEditBtn"), "")
			End If
		End If
	Case "C"
		strCat = getMid(strContent, "startLoop", "endLoop")
		strCat = Replace(strCat, "{DtxtProm}", getsearchCartLngStr("DtxtProm"))
		strCat = Replace(strCat, "{LtxtExpires}", getsearchCartLngStr("LtxtExpires"))
		If Not (myApp.ShowPriceTax and userType = "C" and Session("RetVal") <> "") Then
			strCat = Replace(strCat, "{DtxtPrice}", getsearchCartLngStr("DtxtPrice"))
		Else
			strCat = Replace(strCat, "{DtxtPrice}", getsearchCartLngStr("DtxtPrice") & "+" & txtTax)
		End If

		strCat = Replace(strCat, "{DtxtTemplate}", getsearchCartLngStr("DtxtRecommendations"))
		strCat = Replace(strCat, "{CatCols}", catCols)
		If userType = "C" and InStr(strCat, "startItemTemplate") <> 0 Then
			strCat = Replace(strCat, getFullMid(strCat, "startItemTemplate", "endItemTemplate"), "")
		End If
		itemIndex = 0
		
		If InStr(strCat, "startEditBtn") <> 0 Then
			If myAut.HasAuthorization(170) and searchCmd = "searchCatalog" and Request("PrintCatalog") <> "Y" Then
				strCat = Replace(strCat, getFullMid(strCat, "startEditBtn", "endEditBtn"), getMid(strCat, "startEditBtn", "endEditBtn"))
			Else
				strCat = Replace(strCat, getFullMid(strCat, "startEditBtn", "endEditBtn"), "")
			End If
		End If

	Case "L"
		listColSpan = 3
		
		If userType = "C" Then
			If InStr(strContent, "startItemTemplate") <> 0 Then
				strContent = Replace(strContent, getFullMid(strContent, "startItemTemplate", "endItemTemplate"), "")
			End If
			If InStr(strContent, "startIconFld2") <> 0 Then
				strContent = Replace(strContent, getFullMid(strContent, "startIconFld1", "endIconFld1"), "")
				strContent = Replace(strContent, getFullMid(strContent, "startIconFld2", "endIconFld2"), "")
				listColSpan = listColSpan - 1
			End If
		End If
		
		strList = getMid(strContent, "startLoop", "endLoop")
		strList = Replace(strList, "{DtxtTemplate}", getsearchCartLngStr("DtxtRecommendations"))
		
		If InStr(strList, "startEditBtn") <> 0 Then
			If myAut.HasAuthorization(170) and searchCmd = "searchCatalog" and Request("PrintCatalog") <> "Y" Then
				strList = Replace(strList, getFullMid(strList, "startEditBtn", "endEditBtn"), getMid(strList, "startEditBtn", "endEditBtn"))
			Else
				strList = Replace(strList, getFullMid(strList, "startEditBtn", "endEditBtn"), "")
			End If
		End If

		
End Select

If Not rs.eof Then 
	Select Case CatType
		Case "C"
			Response.Write getMid(strContent, "endPaging", "startLoop")
		Case "L"
			strTop = getMid(strContent, "endPaging", "startLoop")
			If Not (myApp.ShowPriceTax and userType = "C" and Session("RetVal") <> "") Then
				strTop = Replace(strTop, "{DtxtPrice}", getsearchCartLngStr("DtxtPrice"))
			Else
				strTop = Replace(strTop, "{DtxtPrice}", getsearchCartLngStr("DtxtPrice") & "+" & txtTax)
			End If
			strTop = Replace(strTop, "{DtxtQuantity}", getsearchCartLngStr("DtxtQty"))
			If Session("ShowPrice") Then
				strTop = Replace(strTop, getFullMid(strTop, "startShowPrice1", "endShowPrice1"), getMid(strTop, "startShowPrice1", "endShowPrice1"))
			Else
				strTop = Replace(strTop, getFullMid(strTop, "startShowPrice1", "endShowPrice1"), "")
			End If
			
			
			If Not rx.Eof Then
				tmpStr = ""
				
				strTitle = getMid(strTop, "startCustFld1", "endCustFld1")
				
				do while not rx.eof
					ColName = rx("ColName")
					tmpStr = tmpStr & Replace(strTitle, "{Title}", ColName)
					
					listColSpan = listColSpan + 1
				rx.movenext
				loop
				If rx.recordcount > 0 Then rx.movefirst
				
				strTop = Replace(strTop, getFullMid(strTop, "startCustFld1", "endCustFld1"), tmpStr)
			Else
				strTop = Replace(strTop, getFullMid(strTop, "startCustFld1", "endCustFld1"), "")
			End If
			
				
			If Request("PrintCatalog") <> "Y" and Request("excell") <> "Y" Then
				If Request("cmd") <> "searchCatalog" then
					If (optBasket or userType = "V") and Session("username") <> "-Anon-" Then
						strTop = Replace(strTop, getFullMid(strTop, "startBuyBtn1", "endBuyBtn1"), getMid(strTop, "startBuyBtn1", "endBuyBtn1"))
						listColSpan = listColSpan + 1
					Else
						strTop = Replace(strTop, getFullMid(strTop, "startBuyBtn1", "endBuyBtn1"), "")
					End If
				Else
						strTop = Replace(strTop, getFullMid(strTop, "startBuyBtn1", "endBuyBtn1"), "")
				End If
			Else
				strTop = Replace(strTop, getFullMid(strTop, "startBuyBtn1", "endBuyBtn1"), "")
			End If
			
			strList = Replace(strList, "{ColSpan}", listColSpan)
			
			Response.Write strTop
	End Select

	do while not rs.EOF 
		If Request("sourceDoc") <> "" and Request("CPList") = "X" Then
			If rs("UseBaseUn") = "N" Then SaleType = 2 Else SaleType = 1
		End If
		
		If rs("PicturName") <> "" Then
			Pic = olkimgpath & rs("PicturName")
			PicT = rs("PicturName")
		Else
			Pic = olkimgpath & "n_a.gif"
			PicT = "n_a.gif"
		End If
		
		Select Case SaleType
			Case 1
			
				If Session("ShowPrice") Then 
					Price = CDbl(rs("Price"))
					If Not IsNull(rs("DisPrice")) Then DisPrice = CDbl(rs("DisPrice")) Else DisPrice = rs("DisPrice")
				End If
				
				UnPrice = "Un."
				SaleUn = "Un.(1)"
			Case 2
				If Session("ShowPrice") Then
					If Request("sourceDoc") = "" Then 
						Price = CDbl(rs("Price"))*CDbl(rs("NumInSale"))
						If Not IsNull(rs("DisPrice")) Then DisPrice = CDbl(rs("DisPrice"))*CDbl(rs("NumInSale")) Else DisPrice = rs("DisPrice")
					Else
						Price = CDbl(rs("Price"))
						If Not IsNull(rs("DisPrice")) Then DisPrice = CDbl(rs("DisPrice")) Else DisPrice = rs("DisPrice")
					End If
				End If
				
				UnPrice = rs("SalUnitMsr")
				SaleUn = rs("SalUnitMsr") 
				If myApp.GetShowQtyInUn Then SaleUn = SaleUn & "(" & rs("NumInSale") & ")"
				
			Case 3
			
				SaleUn = rs("SalPackMsr")
				If myApp.GetShowQtyInUn Then SaleUn = SaleUn & "(" & rs("SalPackUn") & ")"
				If myApp.UnEmbPriceSet Then
					If Session("ShowPrice") Then 
						Price = CDbl(rs("Price"))*CDbl(rs("NumInSale"))
						If Not IsNull(rs("DisPrice")) Then DisPrice = CDbl(rs("DisPrice"))*CDbl(rs("NumInSale")) Else DisPrice = rs("DisPrice")
					End If
					
					UnPrice = rs("SalUnitMsr")
					SaleUn = SaleUn & " x " & rs("SalUnitMsr")
					If myApp.GetShowQtyInUn Then SaleUn = SaleUn & "(" & rs("NumInSale") & ")"
				Else
					If Session("ShowPrice") Then 
						Price = CDbl(rs("Price"))*CDbl(rs("NumInSale"))*CDbl(rs("SalPackUn"))
						If Not IsNull(rs("DisPrice")) Then DisPrice = CDbl(rs("DisPrice"))*CDbl(rs("NumInSale"))*CDbl(rs("SalPackUn")) Else DisPrice = rs("DisPrice")
					End If
					
					UnPrice = rs("SalPackMsr")
				End If
		End Select
		
		If Session("ShowPrice") and myApp.ShowPriceTax and userType = "C" and Session("RetVal") <> "" Then
			Price = Price + (Price * CDbl(rs("AddVAT")))
			If Not IsNull(rs("DisPrice")) Then DisPrice = DisPrice + (DisPrice * CDbl(rs("AddVAT")))
		End If
		
		If Request("PrintCatalog") <> "Y" and Request("excell") <> "Y" and Request("cmd") <> "searchCatalog" Then If rs("ChkWL") = "N" Then wlGray = "_gray" Else wlGray = ""
		
		Select Case CatType
			Case "L"
				tmpStr = strList
				tmpStr = Replace(tmpStr, "{ItemCode}", Server.HTMLEncode(rs("ItemCode")))
				tmpStr = Replace(tmpStr, "{TreeType}", rs("TreeType"))
				tmpStr = Replace(tmpStr, "{NumInSale}", rs("NumInSale"))
				tmpStr = Replace(tmpStr, "{Qty}", FormatNumber(1, myApp.QtyDec))
				tmpStr = Replace(tmpStr, "{Bookmark}", rs.bookmark)
				
				If myApp.GetShowRef Then
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startShowRef", "endShowRef"), getMid(tmpStr, "startShowRef", "endShowRef"))
				Else
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startShowRef", "endShowRef"), "")
				End If
				tmpStr = Replace(tmpStr, "{ItemName}", rs("ItemName"))
				
				If Not IsNull(rs("PicturName")) Then
					tmpStr = Replace(tmpStr, "{Picture}", rs("PicturName"))
					If InStr(tmpStr, "startPicLink") <> 0 Then 
						tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startPicLink", "endPicLink"), getMid(tmpStr, "startPicLink", "endPicLink"))
						If rs("MoreImg") = "Y" Then
							tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startPicMore", "endPicMore"), getMid(tmpStr, "startPicMore", "endPicMore"))
							If PicT <> "n_a.gif" and Request("Excell") <> "Y" and Request("PrintCatalog") <> "Y" then
								tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startPicMoreLink", "endPicMoreLink"), getMid(tmpStr, "startPicMoreLink", "endPicMoreLink"))
							Else
								tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startPicMoreLink", "endPicMoreLink"), "")
							End If
						Else
							tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startPicMore", "endPicMore"), "")
						End If
					End If
				Else
					If InStr(tmpStr, "startPicLink") <> 0 Then 
						tmpStr = Replace(tmpStr, "{Picture}", "n_a.gif")
						tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startPicLink", "endPicLink"), "")
						tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startPicMore", "endPicMore"), "")
					End If
				End If
				
				If Session("ShowPrice") Then
					tmpStr = Replace(tmpStr, "{Currency}", rs("Currency"))
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startShowPrice2", "endShowPrice2"), getMid(tmpStr, "startShowPrice2", "endShowPrice2"))
					
					If myApp.GetShowSalUn Then
						tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startShowUn", "endShowUn"), getMid(tmpStr, "startShowUn", "endShowUn"))
						tmpStr = Replace(tmpStr, "{UnPrice}", UnPrice)
					Else
						tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startShowUn", "endShowUn"), "")
					End If
					
					If Not IsNull(DisPrice) Then
						tmpStr = Replace(tmpStr, "{Price}", FormatNumber(DisPrice,myApp.PriceDec))
					Else
						If Price <> "" Then tmpStr = Replace(tmpStr, "{Price}", FormatNumber(Price,myApp.PriceDec))
					End If
				Else
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startShowPrice2", "endShowPrice2"), "")
				End If
								
				If Request("PrintCatalog") <> "Y" and Request("excell") <> "Y" Then
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startViewBtn", "endViewBtn"), getMid(tmpStr, "startViewBtn", "endViewBtn"))	
					If Request("cmd") <> "searchCatalog" then
						If (optWish or userType = "V" and myAut.HasAuthorization(26)) and Request("chkWL") <> "Y" and Session("username") <> "-Anon-" Then
							tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startWLBtn", "endWLBtn"), getMid(tmpStr, "startWLBtn", "endWLBtn"))
							tmpStr = Replace(tmpStr, "{IsWL}", wlGray)
						Else
							tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startWLBtn", "endWLBtn"), "")			
						End If
						If (optBasket or userType = "V") and Session("username") <> "-Anon-" Then
							tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startBuyBtn2", "endBuyBtn2"), getMid(tmpStr, "startBuyBtn2", "endBuyBtn2"))
							If InStr(tmpStr, "startBuyChk") <> 0 Then
								tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startBuyChk", "endBuyChk"), getMid(tmpStr, "startBuyChk", "endBuyChk"))
								If InStr(Request("Items"), "'" & rs("ItemCode") & "'") <> 0 Then
									tmpStr = Replace(tmpStr, "{IsChecked}", "checked")
								Else
									tmpStr = Replace(tmpStr, "{IsChecked}", "")
								End If
							End If			
						Else
							tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startBuyBtn2", "endBuyBtn2"), "")
							If InStr(tmpStr, "startBuyChk") <> 0 Then
								tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startBuyChk", "endBuyChk"), "")
							End If			
						End If
					Else
						tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startWLBtn", "endWLBtn"), "")
						tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startBuyBtn2", "endBuyBtn2"), "")
						If InStr(tmpStr, "startBuyChk") <> 0 Then
							tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startBuyChk", "endBuyChk"), "")
						End If
					End If
				Else
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startBuyBtn2", "endBuyBtn2"), "")	
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startViewBtn", "endViewBtn"), "")
					If InStr(tmpStr, "startWLBtn") <> 0 Then tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startWLBtn", "endWLBtn"), "")
					If InStr(tmpStr, "startBuyChk") <> 0 Then tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startBuyChk", "endBuyChk"), "")
				End If
				
				If RS("dias") <= myApp.f_creacion then
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startNew", "endNew"), getMid(tmpStr, "startNew", "endNew"))
				Else
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startNew", "endNew"), "")
				End If
				
				If userType = "V" Then
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startItemTypeImg", "endItemTypeImg"), getMid(tmpStr, "startItemTypeImg", "endItemTypeImg"))
					tmpStr = Replace(tmpStr, "{ItemType}", rs("ItemType"))
					tmpStr = Replace(tmpStr, "{ItemTypeAlt}", rs("ItemTypeAlt"))
					
					If rs("TreeType") <> "T" Then
						tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startItemTemplate", "endItemTemplate"), "")
					End If
				Else
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startItemTypeImg", "endItemTypeImg"), "")
				End If
				
				
				If Not rx.Eof Then
					tmpStrCust = ""
					
					strValue = getMid(tmpStr, "startCustFld2", "endCustFld2")
					
					do while not rx.eof
						ColName = rx("ColName")
						If Not IsNull(rs(ColName)) Then 
							tmpStrCust = tmpStrCust & Replace(strValue, "{Value}", rs(ColName))
						Else
							tmpStrCust = tmpStrCust & Replace(strValue, "{Value}", "")
						End If
						
						If Session("rtl") = "" Then
							tmpStrCust = Replace(tmpStrCust, "{ColAlign}", rx("ColAlign"))
						Else
							Select Case CStr(rx("ColAlign"))
								Case "right"
									tmpStrCust = Replace(tmpStrCust, "{ColAlign}", "left")
								Case "left"
									tmpStrCust = Replace(tmpStrCust, "{ColAlign}", "right")
								Case Else
									tmpStrCust = Replace(tmpStrCust, "{ColAlign}", rx("ColAlign"))
							End Select
						End If
					rx.movenext
					loop
					If rx.recordcount > 0 Then rx.movefirst
					
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startCustFld2", "endCustFld2"), tmpStrCust)
				Else
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startCustFld2", "endCustFld2"), "")
				End If
				
				Response.Write tmpStr


			Case "T"
				tmpStr = strStore
				tmpStr = Replace(tmpStr, "{ItemCode}", Server.HTMLEncode(rs("ItemCode")))
				tmpStr = Replace(tmpStr, "{TreeType}", rs("TreeType"))
				tmpStr = Replace(tmpStr, "{tdWidth}", (ImgMaxSize+21))
				tmpStr = Replace(tmpStr, "{tdHeight}", (ImgMaxSize+10))
				tmpStr = Replace(tmpStr, "{NumInSale}", rs("NumInSale"))
				tmpStr = Replace(tmpStr, "{Qty}", FormatNumber(1, myApp.QtyDec))
				tmpStr = Replace(tmpStr, "{Bookmark}", rs.bookmark)
				
				If myApp.GetShowRef Then
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startShowRef", "endShowRef"), getMid(tmpStr, "startShowRef", "endShowRef"))
					tmpStr = Replace(tmpStr, "{ItemNameSpan}", "")
				Else
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startShowRef", "endShowRef"), "")
					tmpStr = Replace(tmpStr, "{ItemNameSpan}", "colspan=""2""")
				End If
				tmpStr = Replace(tmpStr, "{ItemName}", rs("ItemName"))
				tmpStr = Replace(tmpStr, "{UserText}", myHTMLDecode(rs("UserText")))
				If Not IsNull(rs("PicturName")) Then
					tmpStr = Replace(tmpStr, "{Picture}", rs("PicturName"))
					If InStr(tmpStr, "startPicLink") <> 0 Then 
						tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startPicLink", "endPicLink"), getMid(tmpStr, "startPicLink", "endPicLink"))
						If rs("MoreImg") = "Y" Then
							tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startPicMore", "endPicMore"), getMid(tmpStr, "startPicMore", "endPicMore"))
							If PicT <> "n_a.gif" and Request("Excell") <> "Y" and Request("PrintCatalog") <> "Y" then
								tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startPicMoreLink", "endPicMoreLink"), getMid(tmpStr, "startPicMoreLink", "endPicMoreLink"))
							Else
								tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startPicMoreLink", "endPicMoreLink"), "")
							End If
						Else
							tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startPicMore", "endPicMore"), "")
						End If
					End If
				Else
					If InStr(tmpStr, "startPicLink") <> 0 Then 
						tmpStr = Replace(tmpStr, "{Picture}", "n_a.gif")
						tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startPicLink", "endPicLink"), "")
						tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startPicMore", "endPicMore"), "")
					End If
				End If
				If Session("ShowPrice") Then
					tmpStr = Replace(tmpStr, "{WithoutPList}", rs("WithoutPList"))
					tmpStr = Replace(tmpStr, "{Currency}", rs("Currency"))
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startShowPrice1", "endShowPrice1"), getMid(tmpStr, "startShowPrice1", "endShowPrice1"))
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startShowPrice2", "endShowPrice2"), getMid(tmpStr, "startShowPrice2", "endShowPrice2"))
					If myApp.GetShowSalUn Then
						tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startShowUn", "endShowUn"), getMid(tmpStr, "startShowUn", "endShowUn"))
						tmpStr = Replace(tmpStr, "{UnPrice}", UnPrice)
					Else
						tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startShowUn", "endShowUn"), "")
					End If
					If Not IsNull(DisPrice) Then
						If CDbl(DisPrice) < CDbl(Price) Then
							tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startPriceStrike1", "endPriceStrike1"), getMid(tmpStr, "startPriceStrike1", "endPriceStrike1"))
							tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startPriceStrike2", "endPriceStrike2"), getMid(tmpStr, "startPriceStrike2", "endPriceStrike2"))
							tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startPriceDis", "endPriceDis"), getMid(tmpStr, "startPriceDis", "endPriceDis"))
							tmpStr = Replace(tmpStr, "{DisPrice}", FormatNumber(DisPrice,myApp.PriceDec))
							If Not IsNull(rs("ToDate")) Then
								tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startPriceDisExp", "endPriceDisExp"), getMid(tmpStr, "startPriceDisExp", "endPriceDisExp"))
								tmpStr = Replace(tmpStr, "{ToDate}", FormatDate(rs("ToDate"), True))
							Else
								tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startPriceDisExp", "endPriceDisExp"), "")
							End If
						Else
							Price = DisPrice
							tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startPriceStrike1", "endPriceStrike1"), "")
							tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startPriceStrike2", "endPriceStrike2"), "")
							tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startPriceDis", "endPriceDis"), "")
							tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startPriceDisExp", "endPriceDisExp"), "")
						End If
					Else
						tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startPriceStrike1", "endPriceStrike1"), "")
						tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startPriceStrike2", "endPriceStrike2"), "")
						tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startPriceDis", "endPriceDis"), "")
						tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startPriceDisExp", "endPriceDisExp"), "")
					End If
					If Price <> "" Then tmpStr = Replace(tmpStr, "{Price}", FormatNumber(Price,myApp.PriceDec))
				Else
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startShowPrice1", "endShowPrice1"), "")
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startShowPrice2", "endShowPrice2"), "")
				End If
				
				If Request("PrintCatalog") <> "Y" and Request("excell") <> "Y" Then
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startViewBtn", "endViewBtn"), getMid(tmpStr, "startViewBtn", "endViewBtn"))	
					If Request("cmd") <> "searchCatalog" then
						If (optWish or userType = "V" and myAut.HasAuthorization(26)) and Request("chkWL") <> "Y" and Session("username") <> "-Anon-" Then
							tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startWLBtn", "endWLBtn"), getMid(tmpStr, "startWLBtn", "endWLBtn"))
						Else
							tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startWLBtn", "endWLBtn"), "")			
						End If
						If (optBasket or userType = "V") and Session("username") <> "-Anon-" Then
							tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startBuyBtn", "endBuyBtn"), getMid(tmpStr, "startBuyBtn", "endBuyBtn"))
							If InStr(tmpStr, "startBuyChk") <> 0 Then
								tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startBuyChk", "endBuyChk"), getMid(tmpStr, "startBuyChk", "endBuyChk"))
								If InStr(Request("Items"), "'" & rs("ItemCode") & "'") <> 0 Then
									tmpStr = Replace(tmpStr, "{IsChecked}", "checked")
								Else
									tmpStr = Replace(tmpStr, "{IsChecked}", "")
								End If
							End If			
						Else
							tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startBuyBtn", "endBuyBtn"), "")
							If InStr(tmpStr, "startBuyChk") <> 0 Then
								tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startBuyChk", "endBuyChk"), "")
							End If			
						End If
					Else
						tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startWLBtn", "endWLBtn"), "")
						tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startBuyBtn", "endBuyBtn"), "")
						If InStr(tmpStr, "startBuyChk") <> 0 Then
							tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startBuyChk", "endBuyChk"), "")
						End If
					End If
					
					If RS("dias") <= myApp.f_creacion then
						tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startNew", "endNew"), getMid(tmpStr, "startNew", "endNew"))
					Else
						tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startNew", "endNew"), "")
					End If
					
					If userType = "V" Then
						tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startItemTypeImg", "endItemTypeImg"), getMid(tmpStr, "startItemTypeImg", "endItemTypeImg"))
						tmpStr = Replace(tmpStr, "{ItemType}", rs("ItemType"))
						tmpStr = Replace(tmpStr, "{ItemTypeAlt}", rs("ItemTypeAlt"))
						
						If rs("TreeType") <> "T" Then
							tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startItemTemplate", "endItemTemplate"), "")
						End If
					Else
						tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startItemTypeImg", "endItemTypeImg"), "")
					End If
				Else
					'tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startBuyBtn", "endBuyBtn"), "")
					'tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startViewBtn", "endViewBtn"), "")
					'tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startWLBtn", "endWLBtn"), "")
					'If InStr(tmpStr, "startBuyChk") <> 0 Then
					'	tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startBuyChk", "endBuyChk"), "")
					'End If	
				End If
						
				tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startCustFld", "endCustFld"), getMid(tmpStr, "startCustFld", "endCustFld"))
				strCustFld = getMid(tmpStr, "startCustFldField", "endCustFldField")
				strCustSep = getMid(tmpStr, "startCustFldSep", "endCustFldSep")
				strCust = ""
				ColOrdr = 1
				If Not rx.eof then
					
					If InStr(tmpStr, "{CustFldColspan}") = 0 Then doOldSpan = True Else doOldSpan = False
					do while not rx.eof 
						If rx("ColOrdr") > ColOrdr Then strCust = strCust & strCustSep
						ColOrdr = rx("ColOrdr") 
						ColName = rx("ColName")
						tmpStrCust = strCustFld
						If rx("ColCount") = 1 Then
							If Not doOldSpan Then
								tmpStrCust = Replace(tmpStrCust, "{CustFldColspan}", "colspan=""2""")
							Else
								tmpStrCust = Replace(tmpStrCust, getFullMid(tmpStrCust, "startCustFldColspan", "endCustFldColspan"), getMid(tmpStrCust, "startCustFldColspan", "endCustFldColspan"))
							End If
						Else
							If Not doOldSpan Then
								tmpStrCust = Replace(tmpStrCust, "{CustFldColspan}", "")
							Else
								tmpStrCust = Replace(tmpStrCust, getFullMid(tmpStrCust, "startCustFldColspan", "endCustFldColspan"), "")
							End If
						End If
						If Session("rtl") = "" Then
							tmpStrCust = Replace(tmpStrCust, "{ColAlign}", rx("ColAlign"))
						Else
							Select Case CStr(rx("ColAlign"))
								Case "right"
									tmpStrCust = Replace(tmpStrCust, "{ColAlign}", "left")
								Case "left"
										tmpStrCust = Replace(tmpStrCust, "{ColAlign}", "right")
								Case Else
									tmpStrCust = Replace(tmpStrCust, "{ColAlign}", rx("ColAlign"))
							End Select
						End If
						tmpStrCust = Replace(tmpStrCust, "{ColName}", ColName)
						If Not IsNull(rs(ColName)) and rs(ColName) <> "" Then
							tmpStrCust = Replace(tmpStrCust, "{ColValue}", rs(ColName))
						Else
							tmpStrCust = Replace(tmpStrCust, "{ColValue}", "&nbsp;")
						End If
						strCust = strCust & tmpStrCust
					rx.movenext
					loop
					If rx.recordcount > o Then rx.movefirst
				End If
				
				tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startCustFldSep", "endCustFldSep"), "")
				tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startCustFldField", "endCustFldField"), strCust)
				
				Response.Write tmpStr
			
			Case "C"
				tmpStr = strCat
				tmpStr = Replace(tmpStr, "{ItemCode}", Server.HTMLEncode(rs("ItemCode")))
				tmpStr = Replace(tmpStr, "{TreeType}", rs("TreeType"))
				tmpStr = Replace(tmpStr, "{tdWidth}", CInt(100/(catCols+1)))
				tmpStr = Replace(tmpStr, "{tdHeight}", ImgMaxSize+10)
				tmpStr = Replace(tmpStr, "{ItemName}", rs("ItemName"))
				tmpStr = Replace(tmpStr, "{NumInSale}", rs("NumInSale"))
				tmpStr = Replace(tmpStr, "{TreeType}", rs("TreeType"))
				tmpStr = Replace(tmpStr, "{LtxtConfRemWLItem}", Replace(getsearchCartLngStr("LtxtConfRemWLItem"), "{0}", Server.HTMLEncode(rs("ItemCode"))))
				tmpStr = Replace(tmpStr, "{Qty}", FormatNumber(1, myApp.QtyDec))
				tmpStr = Replace(tmpStr, "{Bookmark}", rs.bookmark)
					
				If myApp.GetShowRef Then
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startShowRef", "endShowRef"), getMid(tmpStr, "startShowRef", "endShowRef"))
				Else
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startShowRef", "endShowRef"), "")
				End If
				
				If Not IsNull(rs("PicturName")) Then
					tmpStr = Replace(tmpStr, "{Picture}", rs("PicturName"))
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startPicLink", "endPicLink"), getMid(tmpStr, "startPicLink", "endPicLink"))
					If rs("MoreImg") = "Y" Then
						tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startPicMore", "endPicMore"), getMid(tmpStr, "startPicMore", "endPicMore"))
						If PicT <> "n_a.gif" and Request("Excell") <> "Y" and Request("PrintCatalog") <> "Y" then
							tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startPicMoreLink", "endPicMoreLink"), getMid(tmpStr, "startPicMoreLink", "endPicMoreLink"))
						Else
							tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startPicMoreLink", "endPicMoreLink"), "")
						End If
					Else
						tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startPicMore", "endPicMore"), "")
					End If
				Else
					tmpStr = Replace(tmpStr, "{Picture}", "n_a.gif")
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startPicLink", "endPicLink"), "")
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startPicMore", "endPicMore"), "")
				End If
						
				If Session("ShowPrice") Then
					tmpStr = Replace(tmpStr, "{WithoutPList}", rs("WithoutPList"))
					tmpStr = Replace(tmpStr, "{Currency}", rs("Currency"))
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startShowPrice", "endShowPrice"), getMid(tmpStr, "startShowPrice", "endShowPrice"))
					If myApp.GetShowSalUn Then
						tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startShowUn", "endShowUn"), getMid(tmpStr, "startShowUn", "endShowUn"))
						tmpStr = Replace(tmpStr, "{UnPrice}", UnPrice)
					Else
						tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startShowUn", "endShowUn"), "")
					End If
					If Not IsNull(DisPrice) Then
						If DisPrice < DisPrice Then
							tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startPriceStrike1", "endPriceStrike1"), getMid(tmpStr, "startPriceStrike1", "endPriceStrike1"))
							tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startPriceStrike2", "endPriceStrike2"), getMid(tmpStr, "startPriceStrike2", "endPriceStrike2"))
							tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startPriceDis", "endPriceDis"), getMid(tmpStr, "startPriceDis", "endPriceDis"))
							tmpStr = Replace(tmpStr, "{DisPrice}", FormatNumber(DisPrice,myApp.PriceDec))
							If Not IsNull(rs("ToDate")) Then
								tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startPriceDisExp", "endPriceDisExp"), getMid(tmpStr, "startPriceDisExp", "endPriceDisExp"))
								tmpStr = Replace(tmpStr, "{ToDate}", FormatDate(rs("ToDate"), True))
							Else
								tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startPriceDisExp", "endPriceDisExp"), "")
							End If
						Else
							Price = DisPrice
							tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startPriceStrike1", "endPriceStrike1"), "")
							tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startPriceStrike2", "endPriceStrike2"), "")
							tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startPriceDis", "endPriceDis"), "")
							tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startPriceDisExp", "endPriceDisExp"), "")
						End If
					Else
						tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startPriceStrike1", "endPriceStrike1"), "")
						tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startPriceStrike2", "endPriceStrike2"), "")
						tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startPriceDis", "endPriceDis"), "")
						tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startPriceDisExp", "endPriceDisExp"), "")
					End If
					If Price <> "" Then tmpStr = Replace(tmpStr, "{Price}", FormatNumber(Price,myApp.PriceDec))
				Else
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startShowPrice", "endShowPrice"), "")
				End If
						
				If Request("PrintCatalog") <> "Y" and Request("excell") <> "Y" Then
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startViewBtn", "endViewBtn"), getMid(tmpStr, "startViewBtn", "endViewBtn"))	
					If Request("cmd") <> "searchCatalog" then
						If (optWish or userType = "V" and myAut.HasAuthorization(26)) and Request("chkWL") <> "Y" and Session("username") <> "-Anon-" Then
							tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startWLBtn", "endWLBtn"), getMid(tmpStr, "startWLBtn", "endWLBtn"))
							tmpStr = Replace(tmpStr, "{IsWL}", wlGray)
						Else
							tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startWLBtn", "endWLBtn"), "")			
						End If
						If (optBasket or userType = "V") and Session("username") <> "-Anon-" Then
							tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startBuyBtn", "endBuyBtn"), getMid(tmpStr, "startBuyBtn", "endBuyBtn"))	
							If InStr(tmpStr, "startBuyChk") <> 0 Then
								tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startBuyChk", "endBuyChk"), getMid(tmpStr, "startBuyChk", "endBuyChk"))
								If InStr(Request("Items"), "'" & rs("ItemCode") & "'") <> 0 Then
									tmpStr = Replace(tmpStr, "{IsChecked}", "checked")
								Else
									tmpStr = Replace(tmpStr, "{IsChecked}", "")
								End If
							End If		
						Else
							tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startBuyBtn", "endBuyBtn"), "")	
							If InStr(tmpStr, "startBuyChk") <> 0 Then
								tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startBuyChk", "endBuyChk"), "")
							End If		
						End If
					Else
						tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startBuyBtn", "endBuyBtn"), "")
						tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startWLBtn", "endWLBtn"), "")	
						If InStr(tmpStr, "startBuyChk") <> 0 Then
							tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startBuyChk", "endBuyChk"), "")		
						End If
					End If
				Else
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startBuyBtn", "endWLBtn"), "")
					If InStr(tmpStr, "startBuyChk") <> 0 Then
						tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startBuyChk", "endBuyChk"), "")
					End If		
				End If
				If RS("dias") <= myApp.f_creacion then
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startNew", "endNew"), getMid(tmpStr, "startNew", "endNew"))
				Else
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startNew", "endNew"), "")
				End If
				If userType = "V" Then
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startItemTypeImg", "endItemTypeImg"), getMid(tmpStr, "startItemTypeImg", "endItemTypeImg"))
					tmpStr = Replace(tmpStr, "{ItemType}", rs("ItemType"))
					tmpStr = Replace(tmpStr, "{ItemTypeAlt}", rs("ItemTypeAlt"))
					
					If rs("TreeType") <> "T" Then
						tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startItemTemplate", "endItemTemplate"), "")
					End If
				Else
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startItemTypeImg", "endItemTypeImg"), "")
				End If
				If Request("cmd") = "wish" Then
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startRemWLBtn", "endRemWLBtn"), getMid(tmpStr, "startRemWLBtn", "endRemWLBtn"))
				Else
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startRemWLBtn", "endRemWLBtn"), "")
				End If
						
				strCustFld = getMid(tmpStr, "startCustFld", "endCustFld")
				strCust = ""
				If Not rx.eof then
					do while not rx.eof 
						ColName = rx("ColName")
						tmpStrCust = strCustFld
						tmpStrCust = Replace(tmpStrCust, "{ColAlign}", rx("ColAlign"))
						tmpStrCust = Replace(tmpStrCust, "{ColName}", ColName)
						If Not IsNull(rs(ColName)) and rs(ColName) <> "" Then
							tmpStrCust = Replace(tmpStrCust, "{ColValue}", rs(ColName))
						Else
							tmpStrCust = Replace(tmpStrCust, "{ColValue}", "&nbsp;")
						End If
						strCust = strCust & tmpStrCust
					rx.movenext
					loop
					If rx.recordcount > o Then rx.movefirst
				End If
				
				tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startCustFld", "endCustFld"), strCust)
				
				If IsNull(DisPrice) Then
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startBlankPriceDis", "endBlankPriceDis"), getMid(tmpStr, "startBlankPriceDis", "endBlankPriceDis"))
					If IsNull(rs("ToDate")) Then
						tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startBlankPriceDisExp", "endBlankPriceDisExp"), getMid(tmpStr, "startBlankPriceDisExp", "endBlankPriceDisExp"))
					Else
						tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startBlankPriceDisExp", "endBlankPriceDisExp"), "")
					End If
				Else
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startBlankPriceDis", "endBlankPriceDisExp"), "")
				End If
						
				Response.Write tmpStr
			End Select

			Select Case CatType
				Case "C"
					If itemIndex = catCols Then
						Response.Write Replace(getMid(strContent, "startSep", "endSep"), "{ColSpan}", (catCols+1))
						itemIndex = -1
					End If
					itemIndex = itemIndex + 1
				Case "L"
						Response.Write Replace(getMid(strContent, "startSep", "endSep"), "{ColSpan}", listColSpan)
			End Select
		rs.MoveNext
		loop
		If catCols <> itemIndex and iPageCurrent = iPageCount Then
			tmpStr = getMid(strContent, "startColsComp", "endColsComp")
			tmpStr = Replace(tmpStr, "{SepWidth}", ((CInt(100/(catCols+1)))*(catCols-itemIndex)))
			tmpStr = Replace(tmpStr, "{SepSpan}", (catCols-itemIndex)+1)
			Response.Write tmpStr
		End If
		
		Select Case CatType
			Case "C"
				Response.Write getMid(strContent, "endLoop", "startSep")
			Case "L"
				Response.Write getMid(strContent, "endSep", "startNoData")
		End Select
		Else %>
		<%=Replace(getMid(strContent, "startNoData", "endNoData"), "{txtNoData}", getsearchCartLngStr("DtxtNoData"))%>
		<% End If %>
		<%=doChkItems%>
		<%=doPagingStr(getMid(strContent, "startPaging", "endPaging"))%>
		<%=Right(strContent, Len(strContent)-InStr(strContent, "<!--endChkAllBtn-->")-18)%>
		</center>
<% 
  If IsNull(Request("DocFlowErr")) or Request("DocFlowErr") = "" and Request("cmd") <> "wish" and strScriptName <> "prom.asp" Then
	  If myApp.AutoSearchOpen = "Y" and rs.recordcount = 1 and Request("page") = "" then 
		rs.movefirst %>
		  <script type="text/javascript">
		  function openResultItem()
		  {
		  	goViewItem('<%=Replace(rs("ItemCode"), "'", "\'")%>');
		  }
		 	window.setTimeout('openResultItem();', 100);
		  </script><% 
	  ElseIf myApp.AutoSearchOpen = "C" and rs.recordcount = 1 and Request("page") = "" and searchCmd <> "searchCatalog" then 
		rs.movefirst
		response.redirect "cart/addCartSubmitM.asp?item=" & rs("ItemCode") & "&T1=1&redir=" & Session("cart") & "&page=1&document=C"
	  End If 
  End If

set rs = Nothing
set rd = Nothing
set rver = Nothing
set rx = Nothing %>
<form name="frmGPage" action="<%=strScriptName%>" method="post">
<% For each itm in Request.Form
If itm <> "page" and itm <> "PrintCatalog" and itm <> "err" and itm <> "errMInv" Then %>
<input type="hidden" name="<%=itm%>" value="<%=myHTMLEncode(Request(itm))%>"><% End If
next %>
<% For each itm in Request.QueryString
If itm <> "page" and itm <> "PrintCatalog" and itm <> "CPList1" and itm <> "submit" and itm <> "err" and itm <> "errMInv" Then %>
<input type="hidden" name="<%=itm%>" value="<%=myHTMLEncode(Request(itm))%>"><% End If
next %>
<input type="hidden" name="page" value="">
<input type="hidden" name="PrintCatalog" value="">
</form>
<form name="frmGoItem" action="item.asp" method="post">
<input type="hidden" name="Item" value="">
<input type="hidden" name="T1" value="">
<input type="hidden" name="cmd" value="">
<input type="hidden" name="pop" value="Y">
<input type="hidden" name="AddPath" value="">
</form>
<%
retVal = ""
isLoadRec = Request("loadRec") <> ""
For each itm in Request.Form
	If itm <> "err" and itm <> "tItem" and itm <> "cmd" and itm <> "retVal" and itm <> "retURL" and itm <> "Item" and itm <> "DocFlowErr" and itm <> "Items" and _
		(not isLoadRec or isLoadRec and itm <> "loadRec" and itm <> "Qty" and itm <> "Price" and itm <> "SaleType" and itm <> "ItmEntry" and itm <> "RecType") Then
		retVal = retVal & "{y}" & itm & "{i}" & Request(itm)
	End If
Next
For each itm in Request.QueryString
	If itm <> "err" and itm <> "tItem" and itm <> "cmd" and itm <> "retVal" and itm <> "retURL" and itm <> "Item" and itm <> "DocFlowErr" and itm <> "Items" and _
		(not isLoadRec or isLoadRec and itm <> "loadRec" and itm <> "Qty" and itm <> "Price" and itm <> "SaleType" and itm <> "ItmEntry" and itm <> "RecType") Then
		retVal = retVal & "{y}" & itm & "{i}" & Request(itm)
	End If
Next 
If Request("cmd") = "searchCashCart" Then 
	redir = "searchCashCart" 
ElseIf Request("cmd") = "wish" Then
	redir = "wish"
Else 
	redir = "no"
End If
%>
<form name="frmGoAddItem" action="cart/addCartSubmitMulti.asp" method="post">
<input type="hidden" name="Item" value="">
<input type="hidden" name="AddPath" value="">
<input type="hidden" name="T1" value="1">
<input type="hidden" name="WithoutPList" value="">
<input type="hidden" name="redir" value="<%=redir%>">
<input type="hidden" name="retVal" value="<%=retVal%>">
<input type="hidden" name="DocConf" value="">
</form>
<% 


Function doChkItems()
	strChkAll = ""
	If Request("PrintCatalog") <> "Y" and Request("excell") <> "Y" and Request("cmd") <> "searchCatalog" and _
		(optBasket or userType = "V") and Session("username") <> "-Anon-" and iPageCount > 0 Then
		If InStr(strContent, "startChkAllBtn") <> 0 Then
			strChkAll = getMid(strContent, "startChkAllBtn", "endChkAllBtn")
			strChkAll = Replace(strChkAll, "{txtChkAll}", getsearchCartLngStr("LtxtChkAll"))
			strChkAll = Replace(strChkAll, "{LtxtBuyChkItms}", getsearchCartLngStr("LtxtBuyChkItms"))
			If Session("rtl") = "" Then
				strChkAll = Replace(strChkAll, "{rtl}", "")
			Else
				strChkAll = Replace(strChkAll, "{rtl}", "Rtl")
			End If
			If CatType = "C" Then
				strChkAll = Replace(strChkAll, "{ColSpan}", (catCols+1))
			End If
		End If		
	End If
	doChkItems = strChkAll
End Function

Function doSearchTtl(str)
	Select Case CatType 
		Case "T" 
			txtTitle = getsearchCartLngStr("DtxtStore")
		Case "L"
			txtTitle = getsearchCartLngStr("DtxtList")
		Case Else
			If Request("cmd") = "wish" Then 
				txtTitle = getsearchCartLngStr("LttlWishList")
			ElseIf strScriptName = "prom.asp" Then 
				txtTitle = txtProms
			Else
				txtTitle = getsearchCartLngStr("DtxtCat")
			End If
	End Select
	str = Replace(str, "{txtTitle}", txtTitle)
	If CatType = "C" Then str = Replace(str, "{ColSpan}", catCols+1)
	str = Replace(str, "{txtSearchRes}", Replace(getsearchCartLngStr("LtxtSearchRetVal"), "{0}", iSearchRecCount))
	doSearchTtl = str
End Function

Function getFrom
	sqlFrom = "from OITM "
	
	If myApp.EnableSearchAlterCode Then
		sqlFrom = sqlFrom & "left outer join OSCN on OSCN.ItemCode = OITM.ItemCode and OSCN.CardCode = N'" & saveHTMLDecode(Session("username"), False) & "' "
	End If
	
	sqlFrom = sqlFrom & "left outer join OITB on OITB.ItmsGrpCod = OITM.ItmsGrpCod "
	If Request("string") <> "" and rdSearchAs = "S" or Request("ItmsGrpNamFrom") <> "" or Request("ItmsGrpNamTo") <> "" Then
		sqlFrom = sqlFrom & "left outer join OMLT G0 on G0.TableName = 'OITB' and G0.FieldAlias = 'ItmsGrpNam' and G0.PK = OITM.ItmsGrpCod " & _  
							"left outer join MLT1 G1 on G1.TranEntry = G0.TranEntry and G1.LangCode = " & SBOLangID & " "
	End If
	sqlFrom = sqlFrom & "left outer join OMRC on OMRC.FirmCode = OITM.FirmCode "
	If Request("string") <> "" and rdSearchAs = "S" or Request("FirmNameFrom") <> "" or Request("FirmNameTo") <> "" Then
		sqlFrom = sqlFrom & "left outer join OMLT F0 on F0.TableName = 'OMRC' and F0.FieldAlias = 'FirmName' and F0.PK = OITM.FirmCode " & _  
							"left outer join MLT1 F1 on F1.TranEntry = F0.TranEntry and F1.LangCode = " & SBOLangID & " "
	End If
	
	sqlFrom = sqlFrom & "left outer join OMLT F2 on F2.TableName = 'OITM' and F2.FieldAlias = 'ItemName' and F2.PK = OITM.ItemCode " & _  
						"left outer join MLT1 F3 on F3.TranEntry = F2.TranEntry and F3.LangCode = " & SBOLangID & " "

	  If (searchCmd <> "searchCatalog" or searchCmd = "searchCatalog" and Request("CPList") <> "") and (optProm or userType = "V" and Request("CPList") <> "X") then
		  sqlFrom = sqlFrom & _
			"left outer join (select X0.ItemCode, X0.Currency, " & _
			"Case X0.AutoUpdt When 'N' Then X0.DisPrice When 'Y' Then X0.ItmPrice-((X0.ItmPrice*X0.Discount)/100) End DisPrice,  " & _
			"X0.ItmPrice BefPrice, X0.FromDate, X0.ToDate, X0.DisPList " & _
			"from " & _
			"( " & _
			"  select P0.ItemCode, " & _
			"  Case P0.Expand When 'N' Then P0.AutoUpdt When 'Y' Then P1.AutoUpdt End AutoUpdt, " & _
			"  Case P0.Expand When 'N' Then P0.Discount When 'Y' Then P1.Discount End Discount, " & _
			"  Case P0.Expand When 'N' Then P0.Price When 'Y' Then P1.Price End DisPrice, " & _
			"  P2.Price ItmPrice, FromDate, ToDate, " & _
			"  Case P0.Expand When 'N' Then P0.Currency Else P1.Currency End Currency, Case P0.Expand When 'N' Then P0.ListNum Else P1.ListNum End DisPList " & _
			"  from OSPP P0 " & _
			"  left outer join SPP1 P1 on P1.ItemCode = P0.ItemCode and P1.CardCode = P0.CardCode " & _
			"  left outer join ITM1 P2 on P2.ItemCode = P0.ItemCode and P2.PriceList = Case P0.Expand When 'N' Then P0.ListNum Else P1.ListNum End " & _
			"  inner join OITM P3 on P3.ItemCode = P0.ItemCode " & _
			"  where P0.CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "' and " & _
			"  (P1.FromDate is null or DateDiff(day,getdate(),P1.FromDate) <= 0) and " & _
			"  (P1.ToDate is null or DateDiff(day,getdate(),P1.ToDate) >= 0) " & _
			"  union " & _
			"  select P0.ItemCode, 'Y' AutoUpdt, P1.Discount, null DisPrice, P2.Price ItmPrice, null FromDate, null ToDate, P2.Currency, @PriceList DisPList " & _
			"  from OITM P0 " & _
			"  inner join OSPG P1 on P1.CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "' and " & _
			"  ( " & _
			"    (P1.ObjType = 52 and P0.ItmsGrpCod = P1.ObjKey) or " & _
			"    (P1.ObjType = 43 and P0.FirmCode = P1.ObjKey) " & _
			"  ) " & _
			"  inner join ITM1 P2 on P2.ItemCode = P0.ItemCode and P2.PriceList = @PriceList " & _
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
			"    P1.AutoUpdt, P1.Discount, P1.Price DisPrice, P2.Price ItmPrice, FromDate, ToDate, " & _
			"    P1.Currency Currency, P1.ListNum DisPList " & _
			"  from OITM P0 " & _
			"  inner join SPP1 P1 on P1.ItemCode = P0.ItemCode " & _
			"  inner join ITM1 P2 on P2.ItemCode = P0.ItemCode and P2.PriceList = P1.ListNum " & _
			"  where P1.CardCode = N'*@PriceList' " & _
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

		  If searchCmd = "searchCatalog" and Request("sourceDoc") <> "" then ' and Request("CPList") = "X"
			If Request("CPList") = "" or Request("CPList") = "X" Then
				If Request("sourceDoc") <> "-4" Then
					sqlFrom = sqlFrom & "inner join " & oTable1 & " T1 on T1.ItemCode = OITM.ItemCode " & _
										"inner join " & oTable & " T0 on T0.docentry = T1.docentry "
				Else
					sqlFrom = sqlFrom & "inner join R3_ObsCommon..DOC1 T1 on T1.ItemCode = OITM.ItemCode collate database_default " & _
										"inner join R3_ObsCommon..TDOC T0 on T0.LogNum = T1.LogNum "
				End If
			Else
				sqlFrom = sqlFrom & "inner join itm1 T1 on T1.ItemCode = OITM.ItemCode and PriceList = " & PriceList & " "
			End If
		  ElseIf searchCmd = "searchCatalog" and Request("CPList") = "" then
			
		  Else
			sqlFrom = sqlFrom & "inner join itm1 T1 on T1.ItemCode = OITM.ItemCode and PriceList = " & PriceList & " "
			If Session("UserName") <> "" Then 
				sqlFrom = sqlFrom & "left outer join OLKWL W0 on W0.CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "' and W0.ItemCode = OITM.ItemCode "
			End If
		  End If
		  
		  If myApp.GetMinInvBy = "W" or CustomInvVer Then
		  	sqlFrom = sqlFrom & "inner join OITW OITW on OITW.ItemCode = OITM.ItemCode and OITW.WhsCode = " & GetWhsCode("OITM") & " "
		  End If
		  
		  sqlFrom = sqlFrom & "left outer join OITT on OITT.Code = OITM.ItemCode "
		  
		  If myApp.ShowPriceTax and userType = "C" and Session("RetVal") <> "" Then
			sqlFrom = sqlFrom & "inner join R3_ObsCommon..TDOC PT4 on PT4.LogNum = " & Session("RetVal") & " " & _  
			"inner join OCRD PT6 on PT6.CardCode = PT4.CardCode collate database_default "
			
			Select Case myApp.LawsSet
				Case "MX", "GT", "CR", "CL", "US", "CA"
					sqlFrom = sqlFrom & "left outer join CRD1 PT5 on PT5.CardCode = PT4.CardCode collate database_default and PT5.AdresType = 'S' and PT5.[Address] = PT4.ShipToCode collate database_default " & _  
										"left outer join ostc PT0 on PT0.Code = Case IndirctTax When 'Y' Then IsNull(TaxCodeAR,'') Else IsNull((Case When VatLiable = 'Y' or @LawsSet = 'US' Then PT5.TaxCode Else 'Disabled' End),'') End "
			End Select
			
			sqlFrom = sqlFrom & "inner join OCTG PT1 on PT1.GroupNum = PT4.GroupNum " & _  
								"left outer join OCDC PT2 on PT2.Code = PT1.DiscCode  " & _  
								"left outer join CDC1 PT3 on PT3.CdcCode = PT1.DiscCode and PT3.LineId = 0 " 
		  End If
		  
	  
	  If Request("FilterID") <> "" Then
	  	arrFilter = Split(Request("FilterID" & Request("FilterID")), ",")
	  	arrFilterValues = Split(Request("FilterValues"), ",")
		arrTables = Split(Request("TableFilterID" & Request("FilterID")), ", ")
		If Request("TableFilterID" & Request("FilterID")) = "" Then arrTables = Split(", ", ", ")
	  	For f = 0 to UBound(arrFilter)
	  		If arrFilter(f) <> "{QryGroup}" Then
	  			If arrTables(f) <> "" Then
	  				sqlFrom = sqlFrom & " inner join [@" & arrTables(f) & "] F" & f & " on F" & f & ".U_ItemCode = OITM.ItemCode and F" & f & "." & arrFilter(f) & " = N'" & Replace(arrFilterValues(f), "'", "''") & "' "
			  	End If
		  	End If
	  	Next
	  End If
	getFrom = sqlFrom
End Function

Function getLogSearchLineQry(ByVal FilterID, ByVal FilterValue)
	getLogSearchLineQry = "set @LineID = IsNull((select Max(LineID)+1 from OLKLogSearchDetails where UserEntry = @UserEntry and SearchID = @SearchID), 0) " & _
						"insert OLKLogSearchDetails(UserEntry, SearchID, LineID, FilterType, FilterValue) " & _
						"values(@UserEntry, @SearchID, @LineID, " & FilterID & ", N'" & Replace(FilterValue, "'", "''") & "') "
End Function

Function getFilter

  If LogSearch Then
  	sqlLogSearch = ""
  End If
  
  sqlFilter = ""
  
  If Not IsNull(CatalogFilter) and CatalogFilter <> "" and (userType = "C" or userType = "V" and CatalogFilterAgent = "Y") Then
  	sqlFilter = sqlFilter & " and OITM.ItemCode not in (" & CatalogFilter & ") "
  End If
  
  If myApp.GetApplyGenFilter and not IgnoreGeneralFilter Then
  	sqlFilter = sqlFilter & " and OITM.ItemCode not in (" & myApp.GetGenFilter & ") "
  End If
  
  If Request("navIndex") <> "" Then
	sqlFilter = sqlFilter & " and OITM.ItemCode in (" & Replace(Replace(NavFilter, "@CardCode", cardCodeFilterValue), "@SlpCode", Session("vendid")) & ") "
  End If
  
  If Not IsNull(searchTreeFilter) and searchTreeFilter <> "" and (strScriptName <> "prom.asp" and Request("cmd") <> "wish" and not (strScriptName = "search.asp" and Request("navIndex") <> "")) Then
  	sqlFilter = sqlFilter & " and OITM.ItemCode not in (" & searchTreeFilter & ") "
  End If
  
  If (Session("UserName") = "-Anon-" or myApp.EnableAnonCart and myApp.AnonCartClient = Session("UserName")) and not IsNull(myApp.AnonSesFilter) and not ApplyAnonCatFilter Then
  	sqlFilter = sqlFilter & " and OITM.ItemCode not in (" & myApp.AnonSesFilter & ") "

  End If
  
  If Request("grupo") <> "" Then
  	sqlFilter = sqlFilter & " and OITM.ItmsGrpCod = " & Request("grupo") & " "
  	If LogSearch Then sqlLogSearch = sqlLogSearch & getLogSearchLineQry(2, Request("grupo"))
  End If
  
  If Request("marca") <> "" Then 
  	sqlFilter = sqlFilter & " and OITM.FirmCode = " & Request("marca") & " "
  	If LogSearch Then sqlLogSearch = sqlLogSearch & getLogSearchLineQry(3, Request("marca"))
  End If
  
  If Request("FilterID") <> "" Then
  	arrFilter = Split(Request("FilterID" & Request("FilterID")), ",")
  	arrFilterValues = Split(Request("FilterValues"), ",")
	arrTables = Split(Request("TableFilterID" & Request("FilterID")), ", ")
	If Request("TableFilterID" & Request("FilterID")) = "" Then arrTables = Split(", ", ", ")
  	For f = 0 to UBound(arrFilter)
  		If arrFilter(f) <> "{QryGroup}" Then
  			If arrTables(f) = "" Then
		  		If InStr(arrFilter(f), "OITM.") = 0 Then strFilter = "OITM." & arrFilter(f) Else strFilter = arrFilter(f)
		  		sqlFilter = sqlFilter & " and " & strFilter & " = N'" & Replace(arrFilterValues(f), "'", "''") & "' "
		  	End If
	  	Else
	  		sqlFilter = sqlFilter & " and OITM.QryGroup" & arrFilterValues(f) & " = 'Y' "
	  	End If
  	Next
  End If
  
  If Request("ItemCodeFrom") <> "" Then sqlFilter = sqlFilter & " and OITM.ItemCode >= N'" & saveHTMLDecode(Request("ItemCodeFrom"), False) & "' "
  If Request("ItemCodeTo") <> "" Then sqlFilter = sqlFilter & " and OITM.ItemCode <= N'" & saveHTMLDecode(Request("ItemCodeTo"), False) & "' "
  
  If Request("ItmsGrpNamFrom") <> "" Then sqlFilter = sqlFilter & " and Convert(nvarchar(4000),IsNull(G1.Trans, OITB.ItmsGrpNam)) >= N'" & saveHTMLDecode(Request("ItmsGrpNamFrom"), False) & "' "
  If Request("ItmsGrpNamTo") <> "" Then sqlFilter = sqlFilter & " and Convert(nvarchar(4000),IsNull(G1.Trans, OITB.ItmsGrpNam)) <= N'" & saveHTMLDecode(Request("ItmsGrpNamTo"), False) & "' "
	
  If Request("FirmNameFrom") <> "" Then sqlFilter = sqlFilter & " and Convert(nvarchar(4000),IsNull(F1.Trans, OMRC.FirmName)) >= N'" & saveHTMLDecode(Request("FirmNameFrom"), False) & "' "
  If Request("FirmNameTo") <> "" Then sqlFilter = sqlFilter & " and Convert(nvarchar(4000),IsNull(F1.Trans, OMRC.FirmName)) <= N'" & saveHTMLDecode(Request("FirmNameTo"), False) & "' "

  If Request("string") <> "" Then
  	If LogSearch Then sqlLogSearch = sqlLogSearch & getLogSearchLineQry(1, Request("string"))
  	Select Case rdSearchAs
  		Case "S"
		  	sqlSearchFilter = ""
		  	sqlSearchDesign = " (OITM.ItemCode like N'%{0}%' or IsNull(F3.Trans, OITM.ItemName) like N'%{0}%' or frgnName like N'%{0}%' or " & _
			  											"Convert(nvarchar(4000),IsNull(F1.Trans, OMRC.FirmName)) like N'%{0}%' or Convert(nvarchar(4000),IsNull(G1.Trans, OITB.ItmsGrpNam)) like N'%{0}%' "
			  		
			If myApp.EnableSearchAlterCode Then sqlSearchDesign = sqlSearchDesign & "or Substitute like N'%{0}%' "
			
			If Not rQryGroups.Eof Then
				do while not rQryGroups.eof
					qryGroup = rQryGroups(0)
					sqlSearchDesign = sqlSearchDesign & Replace(" or Case When OITM.QryGroup{1} = 'Y' Then N'" & Replace(rQryGroups("Name"), "'", "''") & "' Else '' End like N'%{0}%' ", "{1}", qryGroup)
				rQryGroups.movenext
				loop
				rQryGroups.movefirst
			End If
														
			sqlSearchDesign = sqlSearchDesign & ") "
		
		  	For i = 0 to UBound(arrSearchStr)
		  		If arrSearchStr(i) <> "" Then
			  		If sqlSearchFilter <> "" Then sqlSearchFilter = sqlSearchFilter & " and "
			  		sqlSearchFilter = sqlSearchFilter & Replace(sqlSearchDesign, "{0}", mySearchString(arrSearchStr(i)))
			  	End If
		  	Next
		  	sqlFilter = sqlFilter & " and ((" & sqlSearchFilter & ") or "
		  	
		  	If Not myApp.EnableCodeBarsQry Then
			  	sqlFilter = sqlFilter & "OITM.CodeBars = N'" & saveHTMLDecode(Request("string"), False) & "' "
			Else
			  	sqlFilter = sqlFilter & "OITM.CodeBars = (" & Replace(myApp.CodeBarsQry, "@CodeBars", "N'" & saveHTMLDecode(Request("string"), False) & "'") & ") "
			End If
			  	
			If myApp.EnableSearchItmSupp Then
				sqlFilter = sqlFilter & " or OITM.SuppCatNum = N'" & saveHTMLDecode(Request("string"), False) & "'"
			End If

			
			sqlFilter = sqlFilter & ")"
		Case "E"
			sqlFilter = sqlFilter & " and (OITM.ItemCode = N'" & saveHTMLDecode(Request("string"), False) & "' or OITM.frgnName = N'" & saveHTMLDecode(Request("string"), False) & "' or " & _
						"Convert(nvarchar(max),IsNull(IsNull(F3.Trans, OITM.ItemName), '')) = N'" & saveHTMLDecode(Request("string"), False) & "' or "
			
		  	If Not myApp.EnableCodeBarsQry Then
			  	sqlFilter = sqlFilter & "OITM.CodeBars = N'" & saveHTMLDecode(Request("string"), False) & "' "
			Else
			  	sqlFilter = sqlFilter & "OITM.CodeBars = (" & Replace(myApp.CodeBarsQry, "@CodeBars", "N'" & saveHTMLDecode(Request("string"), False) & "'") & ") "
			End If
			
			sqlFilter = sqlFilter & ") "
	End Select
  End If
	
  If Request("cmd") <> "searchCatalog" or searchCmd = "searchCatalog" and Request("CPList") <> "" Then
  	Select Case SaleType
	  Case 2
		sqlAdd1 = "*IsNull(OITM.NumInSale, 1) "
	  Case 3
		sqlAdd1 = "*(IsNull(OITM.NumInSale, 1)*IsNull(OITM.SalPackUn, 1)) "
	End Select

	If Request("PriceFrom") <> "" Then 
		If LogSearch Then sqlLogSearch = sqlLogSearch & getLogSearchLineQry(4, Request("PriceFrom"))
		If optProm or userType = "V" and Request("CPList") <> "X" Then
			sqlFilter = sqlFilter & " and IsNull(DisPrice, IsNull(Price, 0))"
		Else
			sqlFilter = sqlFilter & " and IsNull(Price, 0)"
		End If
		sqlFilter = sqlFilter & sqlAdd1 & " >= " & getNumeric(CDbl(Request("PriceFrom"))) & " "
	Else
		If optProm or userType = "V" and Request("CPList") <> "X" Then
			sqlFilter = sqlFilter & " and IsNull(DisPrice, IsNull(Price, 0))"
		Else
			sqlFilter = sqlFilter & " and IsNull(Price, 0)"
		End If
		sqlFilter = sqlFilter & sqlAdd1 & " >= " & myApp.MinPrice & " "
	End If
	If Request("PriceTo") <> "" Then 
		If LogSearch Then sqlLogSearch = sqlLogSearch & getLogSearchLineQry(5, Request("PriceTo"))
		If optProm or userType = "V" and Request("CPList") <> "X" Then
			sqlFilter = sqlFilter & " and IsNull(DisPrice, IsNull(Price, 0))"
		Else
			sqlFilter = sqlFilter & " and IsNull(Price, 0)"
		End If
		sqlFilter = sqlFilter & sqlAdd1 & " <= " & getNumeric(CDbl(Request("PriceTo"))) & " "
	End If
  End If
  
  If Request("chkProm") = "Y" Then
  	sqlFilter = sqlFilter & " and DisPrice is not null and DisPrice <> IsNull(Price, 0) and DisPList <> 0 "
  	If LogSearch Then sqlLogSearch = sqlLogSearch & getLogSearchLineQry(9, "Y")
  End If
  
  If Request("chkWL") = "Y" Then 
  	sqlFilter = sqlFilter & " and exists(select 'A' from OLKWL where CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "' and ItemCode = OITM.ItemCode) "
  	If LogSearch Then sqlLogSearch = sqlLogSearch & getLogSearchLineQry(10, "Y")
  End If
	
  If Request("new") = "ON" Then 
  	sqlFilter = sqlFilter & " and datediff(day, OITM.createdate, getdate()) <= " & myApp.f_creacion & " "
  	If LogSearch Then sqlLogSearch = sqlLogSearch & getLogSearchLineQry(8, "Y")
  End If

  If myApp.GetMinInvBy = "S" Then MinInvTbl = "OITM" Else MinInvTbl = "OITW"
  invFilter = ""
  If Request("InvFrom") <> "" Then 
	invFilter = invFilter & " " & MinInvTbl & ".OnHand >= " & Request("InvFrom") & " "
	If LogSearch Then sqlLogSearch = sqlLogSearch & getLogSearchLineQry(6, Request("InvFrom"))
  ElseIf myApp.GetEnableMinInv and (CardType <> "S" or IsNull(CardType)) Then ' and objectID <> 23
	invFilter = invFilter & " " & MinInvTbl & ".OnHand >= " & myApp.GetMinInv & " "
  End If
  If Request("InvTo") <> "" Then 
  	If invFilter <> "" Then invFilter = invFilter & " and"
  	invFilter = invFilter & " OITM.OnHand <= " & getNumeric(CDbl(Request("InvTo"))) & " "
  End If
  
  If invFilter <> "" Then sqlFilter = sqlFilter & "and (" & invFilter & " or InvntItem = 'N') "
  
  If Request("pic") = "ON" Then 
  	sqlFilter = sqlFilter & " and PicturName is not null and RTrim(PicturName) <> ''"
  	If LogSearch Then sqlLogSearch = sqlLogSearch & getLogSearchLineQry(7, "Y")
  End If
  
  If searchCmd = "searchCatalog" and Request("sourceDoc") <> "" then ' and Request("CPList") = "X"
	If Request("CPList") = "" or Request("CPList") = "X" Then
		If Request("sourceDoc") <> "-4" Then
			If Request("LinkRep") <> "Y" Then
				sqlFilter = sqlFilter & "and T0.DocNum = " & Request("DocNum") & " "
			Else
				sqlFilter = sqlFilter & "and T0.DocEntry = " & Request("DocNum") & " "
			End If
		Else
			sqlFilter = sqlFilter & " and T0.LogNum = " & Request("DocNum") & " "
		End If
	Else
		If Request("sourceDoc") <> "-4" Then
			sqlFilter = sqlFilter & " and OITM.ItemCode in " & _
			"(select itemcode from " & oTable1 & " X1 "
			
			If Request("LinkRep") <> "Y" Then
				sqlFilter = sqlFilter & " inner join " & oTable & " X0 on X0.DocEntry = X1.DocEntry where docnum = " & Request("DocNum") 
			Else
				sqlFilter = sqlFilter & " where X1.DocEntry = " & Request("DocNum")
			End If
			
			sqlFilter = sqlFilter & ")"
		Else
			sqlFilter = sqlFilter & " and OITM.ItemCode in " & _
			"(select itemcode collate database_default from R3_ObsCommon..DOC1 X1 inner join " & _
			" R3_ObsCommon..TDOC X0 on X0.LogNum = X1.LogNum where X1.LogNum = " & Request("DocNum") & ")"
		End If
	End If
  ElseIf searchCmd = "searchCatalog" and Request("CPList") = "" then
	sqlFilter = sqlFilter & ""
  Else
		sqlFilter = sqlFilter & ""
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
	"	(/*OITM.FrozenFrom is null or*/ DateDiff(day,OITM.FrozenFrom,getdate()) < 0) " & _
	"	and  " & _
	"	(/*OITM.FrozenTo is null or*/ DateDiff(day,getdate(),OITM.FrozenTo) < 0) " & _
	")) "
	
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
			strQryGroups = strQryGroups & "OITM.QryGroup" & chkQryGroups(i) & " = 'Y'"
		Next
		strQryGroups = strQryGroups & ") "
		
		Select Case Request("QryGroupOp2")
			Case "I"
				sqlFilter = sqlFilter & " and " & strQryGroups
			Case "N"
				sqlFilter = sqlFilter & " and not " & strQryGroups
		End Select
	End If

	
  'Agregar log de filtro de cliente
  If LogSearch and sqlLogSearch <> "" Then
  	sqlLogSearch = "declare @UserEntry int set @UserEntry = (select DocEntry from OCRD where CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "') " & _
		"declare @SearchID int set @SearchID = IsNull((select Max(SearchID)+1 from OLKLogSearch where UserEntry = @UserEntry), 0) " & _
		"declare @LineID int " & _
		"insert OLKLogSearch(UserEntry, SearchID, TimeStamp) " & _
		"values(@UserEntry, @SearchID, getdate()) " & sqlLogSearch
	conn.execute(sqlLogSearch)
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
  	If InStr(Trim(SearchQuery), "@SystemFilters") = 1 Then
		andPos = InStr(LCase(sqlFilter), "and")
		sqlFilter = Mid(sqlFilter, andPos+3, Len(sqlFilter)-andPos-3)
	  	SearchQuery = Replace(SearchQuery, "@SystemFilters", "and (" & sqlFilter & ")")
	Else
		If InStr(Trim(SearchQuery), "@SystemFilters") <> 0 Then
			andPos = InStr(LCase(sqlFilter), "and")
			sqlFilter = Mid(sqlFilter, andPos+3, Len(sqlFilter)-andPos-3)
	  		SearchQuery = "and " & Replace(SearchQuery, "@SystemFilters", "(" & sqlFilter & ")")
		Else
			SearchQuery = "and " & SearchQuery
		End If
	End If
	
	SearchQuery = Replace(SearchQuery, "@SlpCode", Session("vendid"))
	SearchQuery = Replace(SearchQuery, "@branch", Session("branch"))
	SearchQuery = Replace(SearchQuery, "@LanID", Session("LanID"))
	If Session("UserName") <> "" Then 
		SearchQuery = Replace(SearchQuery, "@CardCode", "N'" & saveHTMLDecode(Session("UserName"), False) & "'")
	Else
		SearchQuery = Replace(SearchQuery, "@CardCode", "NULL")
	End If

  	getFilter = SearchQuery
  Else
	  getFilter = sqlFilter
  End If
End Function

Function GetSearchSqlStmt()
	sqlFilter = getFilter
  
	If Request("CPList") <> "" Then
		PriceList = Request("CPList")
	ElseIf Request("Excell") = "Y" Then
		PriceList = 0
	Else
		PriceList = Session("PriceList")
	End If
	
	mySql = ""
	
	If Request("string") <> "" and rdSearchAs = "S" Then
		mySql = "select ItemCode, " & rateStr & " [Rate] from ("
	End If
	
	mySql = mySql & "select OITM.ItemCode "
	
	If Request("string") <> "" and rdSearchAs = "S" Then
		Select Case strOrden1
			Case "ItemName"
				mySql = mySql & ", Convert(nvarchar(max),IsNull(IsNull(F3.Trans, OITM.ItemName), '')) ItemName "
			Case "Price"
				If optProm or userType = "V" and Request("CPList") <> "X" Then
					mySql = mySql & ", IsNull(DisPrice, IsNull(Price, 0)) Price "
				Else
					mySql = mySql & ", IsNull(Price, 0) Price "
				End If
		End Select
		mySql = mySql & " , OITM.ItemCode + ' ' + Convert(nvarchar(max),IsNull(IsNull(F3.Trans, OITM.ItemName), '')) + IsNull(' ' + frgnName, '') + IsNull(' ' + Convert(nvarchar(4000),IsNull(F1.Trans, OMRC.FirmName)), '') + IsNull(' ' + Convert(nvarchar(4000),IsNull(G1.Trans, OITB.ItmsGrpNam)), '') "
		If myApp.EnableSearchAlterCode Then mySql = mySql & "+ IsNull(' ' + Substitute, '') "
			If Not rQryGroups.Eof Then
			do while not rQryGroups.eof
				qryGroup = rQryGroups(0)
				mySql = mySql & Replace(" + Case When OITM.QryGroup{0} = 'Y' Then N' " & Replace(rQryGroups("Name"), "'", "''") & "' Else '' End ", "{0}", qryGroup)
			rQryGroups.movenext
			loop
			rQryGroups.movefirst
		End If
		mySql = mySql & " RateStr "
	End If
	
	mySql = mySql & getFrom
  
	mySql = mySql &	" where OITM.SellItem = 'Y' and OITM.Canceled = 'N' and OITM.TreeType <> 'T' "
  
	If arty <> "QryGroup-1" Then
		mySql = mySql & " and OITM." & arty & " = 'N' "
	End If
  
	mySql = mySql & sqlFilter
	
	If Request("string") <> "" and rdSearchAs = "S" Then
		mySql = mySql & ") OITM order by [Rate] desc, " & strOrden1 & " " & strOrden2
	Else
		mySql = mySql & " order by " & strOrden1 & " " & strOrden2
	End If

	If Request("CPList") <> "" Then
		mySql = Replace(mySql,"@PriceList",Request("CPList"))
	ElseIf Request("Excell") = "Y" Then
		mySql = Replace(mySql,"@PriceList","0")
	Else
		mySql = Replace(mySql,"@PriceList",Session("PriceList"))
	End If
	
	GetSearchSqlStmt = mySql
End Function %><!--#include file="itemDetails.inc"-->
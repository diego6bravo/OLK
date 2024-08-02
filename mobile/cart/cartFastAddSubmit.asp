<!--#include file="../myHTMLEncode.asp"-->
<!--#include file="../lcidReturn.inc" -->
<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<!--#include file="../itemFunctions.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

set rs = Server.CreateObject("ADODB.Recordset")

sql = 	"select Object from R3_ObsCommon..TLOG where LogNum = " & Session("RetVal")
set rs = conn.execute(sql)
IsBarCode = False

If myApp.EnableCodeBarsQry Then CodeBarsQry = myApp.CodeBarsQry
ChkInv = rs("Object") <> 23 

sql = "select (select ItemCode from OITM "


If myApp.EnableCodeBarsQry and myApp.CodeBarsQryMethod = "I" Then
	sql = sql & "cross join (" & Replace(myApp.CodeBarsQry, "@CodeBars", "N'" & saveHTMLDecode(Request("txtFastAdd"), False) & "'") & ") tCodeBars "
End If

sql = sql & "where ItemCode = N'" & saveHTMLDecode(Request("txtFastAdd"), False) & "' or "

If Not myApp.EnableCodeBarsQry Then
	sql = sql & "OITM.CodeBars = N'" & saveHTMLDecode(Request("txtFastAdd"), False) & "'"
Else

	Select Case myApp.CodeBarsQryMethod
		Case "R"
		  	sql = sql & "OITM.CodeBars = (" & Replace(CodeBarsQry, "@CodeBars", "N'" & saveHTMLDecode(Request("txtFastAdd"), False) & "'") & ") "
		Case "I"
			sql = sql & "OITM.CodeBars = tCodeBars.CodeBars "
	End Select

	
End If


If myApp.EnableSearchItmSupp Then
	sql = sql & " or SuppCatNum = N'" & saveHTMLDecode(Request("txtFastAdd"), False) & "'"
End If

sql = sql & ") ItemCode, Case When (select ItemCode from OITM "


If myApp.EnableCodeBarsQry and myApp.CodeBarsQryMethod = "I" Then
	sql = sql & "cross join (" & Replace(myApp.CodeBarsQry, "@CodeBars", "N'" & saveHTMLDecode(Request("txtFastAdd"), False) & "'") & ") tCodeBars "
End If

sql = sql & " where "

If Not myApp.EnableCodeBarsQry Then
	sql = sql & "OITM.CodeBars = N'" & saveHTMLDecode(Request("txtFastAdd"), False) & "'"
Else
	Select Case myApp.CodeBarsQryMethod
		Case "R"
		  	sql = sql & "OITM.CodeBars = (" & Replace(CodeBarsQry, "@CodeBars", "N'" & saveHTMLDecode(Request("txtFastAdd"), False) & "'") & ") "
		Case "I"
			sql = sql & "OITM.CodeBars = tCodeBars.CodeBars "
	End Select
End If

sql = sql & ") is not null Then 'Y' Else 'N' End IsBarCode "

set rs = conn.execute(sql)

If Request("SaleUnit") = "" Then 
	SaleType = myApp.GetSaleUnit 
Else 
	SaleType = Request("SaleUnit")
	If myApp.FastAddUnRem Then
		Session("CurSaleType") = SaleType
	End If
End If

If Not IsNull(rs("ItemCode")) Then
	ItemCode = rs("ItemCode")
	IsBarCode = rs("IsBarCode") = "Y"

		sql = "select CatalogFilter, CatalogFilterAgent from OLKClientsAccess where CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "'"
		set rs = conn.execute(sql)
		If Not rs.Eof Then
			CatalogFilter = rs(0)
			CatalogFilterAgent = rs(1)
		End If


		sql = "(OITM.ValidFor = 'N' or OITM.ValidFor = 'Y' and " & _
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
		
	  If Not IsNull(CatalogFilter) and CatalogFilter <> "" and userType = "V" and CatalogFilterAgent = "Y" Then
	  	sql = sql & " and OITM.ItemCode not in (" & CatalogFilter & ") "
	  End If
	  
	  If myApp.GetApplyGenFilter Then
	  	sql = sql & " and OITM.ItemCode not in (" & myApp.GetGenFilter & ") "
	  End If
	  
	  sql = "declare @CardCode nvarchar(15) set @CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "' " & _
	  		"select ItemCode from OITM where ItemCode = N'" & saveHTMLDecode(ItemCode, False) & "' and " & sql
	  set rs = conn.execute(sql)
	  If rs.Eof Then
	  		Response.Redirect "../operaciones.asp?cmd=cart&fastAddErr=Y&fastAddErrItm=" & saveHTMLDecode(Request("txtFastAdd"), False) & "&fastAddErrType=B"
	  End If
	
	AddErr = getAddItmError(ItemCode)
	If CStr(AddErr) <> "" Then	
		retURL = ""
		For each itm in Request.Form
			If retURL <> "" Then retURL = retURL & "{a}"
			retURL = retURL & itm & "{e}" & Request(itm)
		Next
		For each itm in Request.QueryString
			If retURL <> "" Then retURL = retURL & "{a}"
			retURL = retURL & itm & "{e}" & Request(itm)
		Next
		Response.Redirect "../operaciones.asp?cmd=DocFlowErr&DocFlowErr=" & AddErr & _
			"&FlowItem=" & ItemCode & "&ItemFastAdd=Y&retURL=" & retURL
	End If

	
	If ChkInv Then
		sql = 	"declare @ItemCode nvarchar(20) set @ItemCode = N'" & saveHTMLDecode(ItemCode, False) & "' " & _
				"declare @whscode nvarchar(8) set @whscode = OLKCommon.dbo.DBOLKGetWhsCode" & Session("ID") & "(" & Session("branch") & ", " & Session("vendid") & ", @ItemCode) " & _
				"declare @SaleType2 int set @SaleType2 = " & SaleType & " " & _
				"declare @FirstQuantity numeric(19,6) set @FirstQuantity = " & getNumeric(Request("txtFastAddQty")) & " " & _
				"declare @UnEmbPriceSet char(1) set @UnEmbPriceSet = (select UnEmbPriceSet from olkcommon) " & _
				"declare @Quantity numeric(19,6) " & _
				"If @SaleType2 = 1 Begin  " & _
				"set @Quantity = @FirstQuantity/(select NumInSale from oitm where itemcode = @ItemCode) End  " & _
				"Else If @SaleType2 = 2 or @SaleType2 = 3 and @UnEmbPriceSet = 'Y' Begin  " & _
				"set @Quantity = @FirstQuantity End  " & _
				"Else If @SaleType2 = 3 and @UnEmbPriceSet = 'N' Begin  " & _
				"set @Quantity = @FirstQuantity*(select SalPackUn from oitm where itemcode = @ItemCode) End  " & _
				"select OLKCommon.dbo.DBOLKItemInv" & Session("ID") & "(@ItemCode, @WhsCode, @Quantity, '" & Session("olkdb") & "', -1, -1) Verfy"
		set rs = conn.execute(sql)
		If rs("Verfy") <> "Y" Then 	Response.Redirect "../operaciones.asp?cmd=cart&Item="& Request("txtFastAdd") & "&err=disp"
	End If
	

	sql = 	"declare @ItemCode nvarchar(20) set @ItemCode = N'" & saveHTMLDecode(ItemCode, False) & "' " & _
			"declare @whscode nvarchar(8) set @whscode = OLKCommon.dbo.DBOLKGetWhsCode" & Session("ID") & "(" & Session("branch") & ", " & Session("vendid") & ", @ItemCode) " & _
			"select (select case when exists(select itemcode from r3_obscommon..doc1 where lognum = " & Session("RetVal") & _
			" and ItemCode = @ItemCode and whscode = @whscode) Then 'True' ELSE 'False' END) as Verfy, ISNULL((select Max(LineNum) from r3_obscommon..doc1 " & _
			"where LogNum = " & Session("RetVal") & " and ItemCode = @ItemCode and whscode = @whscode),0) As LineNum, " & _
			"ISNULL((select max(linenum)+1 from r3_obscommon..doc1 where lognum = " & Session("RetVal") & "),0) As MaxLineNum"
	set rs = conn.execute(sql)
	
	'no fue agregado lo agrega
	If rs("Verfy") = "False" or rs("Verfy") = "True" and myApp.BasketMItems Then
		TaxCode = ""
		Select Case myApp.LawsSet
			Case "MX", "CL", "CR", "GT", "US", "CA", "BR"
				If Request("TaxCode") <> "" Then
					TaxCode = Request("TaxCode")
				Else
					TaxCode = getItemTaxCode(ItemCode)
				End If
				
				If TaxCode = "Disabled" Then 
					TaxCode = "NULL"
				End If
		End Select
		
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKCartAddSFM" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@lognum") = Session("RetVal")
		cmd("@quantity") = CDbl(getNumericOut(Request("txtFastAddQty")))
		cmd("@item") = ItemCode
		If IsBarCode Then cmd("@CodeBars") = Request("txtFastAdd")
		cmd("@PriceList") = Session("PList")
		cmd("@UserType") = "V"
		cmd("@SaleType") = SaleType
		cmd("@branchIndex") = Session("branch")
		cmd("@SlpCode") = Session("vendid")
		      
		Select Case myApp.LawsSet 
			Case "MX", "CL", "CR", "GT", "US", "CA", "BR"
			If TaxCode <> "NULL" Then cmd("@TaxCode") = TaxCode
		End Select
	
		cmd.execute() 
	'si ya fue agregado se le suma la cantidad que quiere comprar a la que ya fue agregada.
	ElseIf rs("Verfy") = "True" Then
		sql = ""
		If myApp.GetSaleUnit = 3 Then
			sql = "*(select SalPackUn from oitm where itemcode = N'" & saveHTMLDecode(ItemCode, False) & "') "
		End If
		sql = "update r3_obscommon..doc1 set quantity = quantity + " & getNumeric(Request("txtFastAddQty")) & sql & " where LogNum = " & Session("RetVal") & " and linenum = " & rs("LineNum")
		conn.execute(sql)
	End If

	Response.Redirect "../operaciones.asp?cmd=cart&txtFastAddQty=" & Request("txtFastAddQty")

	
Else
	Response.Redirect "../operaciones.asp?cmd=cart&fastAddErr=Y&fastAddErrItm=" & saveHTMLDecode(Request("txtFastAdd"), False)
End If

Function getAddItmError(Item)
	RetVal = ""

	set rFlow = Server.CreateObject("ADODB.RecordSet")
	set rChk = Server.CreateObject("ADODB.RecordSet")
	
	sqlFlow = 	"declare @ObjectCode int set @ObjectCode = (select Object from R3_ObsCommon..TLOG where LogNum = " & Session("RetVal") & ") " & _
				"select T0.FlowID, T0.Name, Type, Query " & _
				"from OLKUAF T0  " & _
				"inner join OLKUAF1 T1 on T1.FlowID = T0.FlowID and T1.SlpCode in (" & Session("vendid") & ",-999) " & _
				"inner join OLKUAF2 T2 on T2.FlowID = T0.FlowID " & _
				"where T2.ObjectCode = @ObjectCode and T0.Active = 'Y' and T0.ExecAt = 'D2' "
	
	If Request("DocConf") <> "" Then sqlFlow = sqlFlow & " and T0.FlowID not in (" & Request("DocConf") & ") "
	
	sqlFlow = sqlFlow & " order by Type, [Order] asc"
	
	set rFlow = conn.execute(sqlFlow)
	sqlBase = 	"declare @LogNum int set @LogNum = " & Session("RetVal") & " " & _
				"declare @CardCode nvarchar(15) set @CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "' " & _
				"declare @SlpCode int set @SlpCode = " & Session("VendID") & " " & _
				"declare @dbName nvarchar(100) set @dbName = db_name() " & _
				"declare @branch int set @branch = " & Session("branch") & " " & _
				"declare @ItemCode nvarchar(20) set @ItemCode = N'" & saveHTMLDecode(Item, False) & "' " & _
				"declare @WhsCode nvarchar(8) set @WhsCode = OLKCommon.dbo.DBOLKGetWhsCode" & Session("ID") & "(" & Session("branch") & ", " & Session("vendid") & ", @ItemCode) " & _
				"declare @UserType char(1) set @UserType = '" & userType & "' " & _
				"declare @SaleType int set @SaleType = " & SaleType & " " & _
				"declare @Quantity numeric(19,6) set @Quantity = 1 " & _
				"set @Quantity = (select @Quantity*Case @SaleType When 1 Then 1 When 2 Then NumInSale When 3 Then NumInSale*SalPackUn End from OITM where ItemCode = @ItemCode) " & _
				"declare @Unit int set @Unit = @SaleType " & _
				"declare @Price numeric(19,6) " & _
				"EXEC OLKCommon..DBOLKGetItemPrice" & Session("ID") & " @ItemCode = @ItemCode, @CardCode = @CardCode, @PriceList = " & Session("Plist") & ", @UserType = 'V', @ItemPrice = @Price out "
	
	do while not rFlow.eof
		sql = sqlBase & rFlow("Query")
		'response.write sql
		set rChk = conn.execute(sql)
		If not rChk.eof then
			If Not IsNull(rChk(0)) Then
				If lcase(rChk(0)) = lcase("True") Then
					If RetVal <> "" Then RetVal = RetVal & ", "
					RetVal = RetVal & rFlow("FlowID")
					If rFlow("Type") = 0 Then Exit do
				End If
			End If
		End If
	rFlow.movenext
	loop
	getAddItmError = RetVal
End Function

%>
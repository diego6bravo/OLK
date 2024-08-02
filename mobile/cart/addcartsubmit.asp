<%@ Language=VBScript %>
<!--#include file="../myHTMLEncode.asp"-->
<!--#include file="../lcidReturn.inc"-->
<% If session("OLKDB") = "" Then response.redirect "../lock.asp" %>
<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

      set rs = Server.CreateObject("ADODB.recordset")
      sql = "select Object from r3_obscommon..tlog cross join olkcommon where lognum = " & Session("RetVal")
      set rs = conn.execute(sql)
      If rs("Object") <> 23 Then ChkInv = True Else ChkInv = False
      
	AddErr = getAddItmError()
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
			"&FlowItem=" & Request("Item") & "&FlowWhsCode=" & Request("WhsCode") & "&FlowQuantity=" & Request("Quantity") & "&Flowprecio=" & Request("precio") & "&FlowSaleType=" & Request("SaleType2") & "&FlowSellAll=" & Request("chkAddAll") & "&retURL=" & retURL
	End If
      
set rv = Server.CreateObject("ADODB.recordset")
set rd = Server.CreateObject("ADODB.recordset")

'Revisa si hay suficiente Stock para agregarlo al shopping cart
If ChkInv Then
	sql = "declare @whscode nvarchar(8) set @whscode = "
	
	If Request("WhsCode") <> "" Then 
		sql = sql & "N'" & saveHTMLDecode(Request("WhsCode"), False) & "' "
	Else
		sql = sql & "OLKCommon.dbo.DBOLKGetWhsCode" & Session("ID") & "(" & Session("branch") & ", " & Session("vendid") & ", N'" & saveHTMLDecode(Request("Item"), False) & "') "
	End If
	
	sql = sql & "declare @SaleType2 int set @SaleType2 = " & Request("SaleType2") & " declare @ItemCode nvarchar(20) set @ItemCode = N'" & saveHTMLDecode(Request("Item"), False) & "' " & _
		  "declare @NumInSale int declare @SalPackUn int select @NumInSale = NumInSale, @SalPackUn = SalPackUn from OITM where ItemCode = @ItemCode " & _
		  "declare @Quantity numeric(19,6) " & _
		  "declare @UnEmbPriceSet char(1) set @UnEmbPriceSet = (select UnEmbPriceSet from olkcommon) "
	  
	If Request("chkAddAll") <> "Y" Then
		sql = sql & "declare @FirstQuantity numeric(19,6) set @FirstQuantity = " & getNumeric(Request("Quantity")) & " " & _
					  "If @SaleType2 = 1 Begin  " & _
					  "set @Quantity = @FirstQuantity/@NumInSale End  " & _
					  "Else If @SaleType2 = 2 or @SaleType2 = 3 and @UnEmbPriceSet = 'Y' Begin  " & _
					  "set @Quantity = @FirstQuantity End  " & _
					  "Else If @SaleType2 = 3 and @UnEmbPriceSet = 'N' Begin  " & _
					  "set @Quantity = @FirstQuantity*@SalPackUn End  "
	Else
		sql = sql & "set @Quantity = OLKCommon.dbo.DBOLKItemInv" & Session("ID") & "Val(@ItemCode, @WhsCode, '" & Session("olkdb") & "', " & Session("RetVal") & ", -1)/@NumInSale "
	End If
	sql = sql & "select OLKCommon.dbo.DBOLKItemInv" & Session("ID") & "(@ItemCode, @WhsCode, @Quantity, '" & Session("olkdb") & "', -1, -1) Verfy"
Else
	sql = "select 'Y' As Verfy"
End If
set rv = conn.execute(sql)

If rv("Verfy") = "Y" Then
	'Revisa si el articulo ya fue agregado a la canasta de compras
	sql5 = "declare @whscode nvarchar(8) set @whscode = "
	
	If Request("WhsCode") <> "" Then 
		sql5 = sql5 & "N'" & saveHTMLDecode(Request("WhsCode"), False) & "' "
	Else
		sql5 = sql5 & "OLKCommon.dbo.DBOLKGetWhsCode" & Session("ID") & "(" & Session("branch") & ", " & Session("vendid") & ", N'" & saveHTMLDecode(Request("Item"), False) & "') "
	End If
	
	sql5 = sql5 & 	"select (select case when exists(select itemcode from r3_obscommon..doc1 where lognum = " & Session("RetVal") & _
				   " and ItemCode = N'" & saveHTMLDecode(Request("Item"), False) & "' and whscode = @whscode) Then 'True' ELSE 'False' END) as Verfy, ISNULL((select Max(LineNum) from r3_obscommon..doc1 " & _
				   "where LogNum = " & Session("RetVal") & " and ItemCode = N'" & saveHTMLDecode(Request("Item"), False) & "' and whscode = @whscode),0) As LineNum, (select sdkId from r3_obscommon..tcif where CompanyDB = N'" & Session("OLKDb") & "') As SDKID, " & _
				   "ISNULL((select max(linenum)+1 from r3_obscommon..doc1 where lognum = " & Session("RetVal") & "),0) As MaxLineNum"
	set rv = conn.execute(sql5)

	'Si no fue agregado lo agrega
	If RV("Verfy") = "False" or RV("Verfy") = "True" and myApp.BasketMItems Then

		If Request("chkAddAll") = "Y" Then addQty = 1 Else addQty = CDbl(getNumericOut(Request.Form("Quantity")))
		
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKCartAddSF" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@PriceList") = Session("PList")
		cmd("@FirstQuantity") = addQty
		cmd("@FirstPrice") = CDbl(getNumericOut(Request.Form("precio")))
		cmd("@LogNum") = Session("RetVal")
		cmd("@ItemCode") = Request("Item")
		If Request("WhsCode") <> "" Then cmd("@WhsCode") = Request("WhsCode")
		cmd("@SaleType") = Request("SaleType2")
		cmd("@ManPrc") = Request("ManPrc")
		cmd("@branch") = Session("branch")
		If Request("TaxCode") <> "" Then cmd("@TaxCode") = Request("TaxCode")
		If Request("chkAddAll") = "Y" Then cmd("@All") = "Y"
		cmd("@SlpCode") = Session("vendid")
		cmd.execute()
	'si ya fue agregado se le suma la cantidad que quiere comprar a la que ya fue agregada.
	ElseIf RV("Verfy") = "True" Then
		If Request("SaleType2") = 3 Then
			sqlAdd = "*(select SalPackUn from oitm where itemcode = N'" & saveHTMLDecode(Request("Item"), False) & "') "
		End If
		sqladd = "update r3_obscommon..doc1 set quantity = quantity + " & getNumeric(Request("Quantity")) & sqlAdd & _
				 " where lognum = " & Session("RetVal") & " and linenum = " & RV("LineNum")
		conn.execute(sqladd)
	End If

	Conn.Close
	If myApp.AfterCartAddPocket = "Y" Then response.redirect "../operaciones.asp?cmd=slistsearch" Else response.redirect "../operaciones.asp?cmd=cart"
Else
	Response.Redirect "../operaciones.asp?cmd=addcart&Item="&Request("item")&"&PackPrice=" & Request("PackPrice") & "&WhsCode=" & Request("WhsCode") & "&err=disp"
End If

Function getAddItmError()
	RetVal = ""

	set rFlow = Server.CreateObject("ADODB.RecordSet")
	set rChk = Server.CreateObject("ADODB.RecordSet")
	
	sqlFlow = 	"declare @ObjectCode int set @ObjectCode = (select Object from R3_ObsCommon..TLOG where LogNum = " & Session("RetVal") & ") " & _
				"select T0.FlowID, T0.Name, Type, Query " & _
				"from OLKUAF T0  "
	
	If userType = "V" Then
		sqlFlow = sqlFlow & "inner join OLKUAF1 T1 on T1.FlowID = T0.FlowID and T1.SlpCode in (" & Session("vendid") & ",-999) "
	End If
	
	sqlFlow = sqlFlow & "inner join OLKUAF2 T2 on T2.FlowID = T0.FlowID " & _
				"where T2.ObjectCode = @ObjectCode and T0.Active = 'Y' and T0.ExecAt = 'D2' "
	
	If userType = "C" Then 
		sqlFlow = sqlFlow & " and T0.ApplyToClient = 'Y' "
	End If
	
	If Request("DocConf") <> "" Then sqlFlow = sqlFlow & " and T0.FlowID not in (" & Request("DocConf") & ") "
	
	sqlFlow = sqlFlow & " order by Type, [Order] asc"
	'response.redirect "http://www.topmanage.com.pa/query.asp?query=" & sqlFlow
	
	set rFlow = conn.execute(sqlFlow)
	sqlBase = 	"declare @LogNum int set @LogNum = " & Session("RetVal") & " " & _
				"declare @CardCode nvarchar(15) set @CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "' " & _
				"declare @SlpCode int set @SlpCode = " & Session("VendID") & " " & _
				"declare @dbName nvarchar(100) set @dbName = db_name() " & _
				"declare @branch int set @branch = " & Session("branch") & " " & _
				"declare @ItemCode nvarchar(20) set @ItemCode = N'" & saveHTMLDecode(Request("Item"), False) & "' " & _
				"declare @WhsCode nvarchar(8) set @WhsCode = "
				
	If Request("WhsCode") <> "" Then 
		sqlBase = sqlBase & "N'" & saveHTMLDecode(Request("WhsCode"), False) & "' "
	Else
		sqlBase = sqlBase & "OLKCommon.dbo.DBOLKGetWhsCode" & Session("ID") & "(" & Session("branch") & ", " & Session("vendid") & ", N'" & saveHTMLDecode(Request("Item"), False) & "') "
	End If

	sqlBase = sqlBase & "declare @UserType char(1) set @UserType = '" & userType & "' " & _
				"declare @SaleType int set @SaleType = " & Request("SaleType2") & " " & _
				"declare @Quantity numeric(19,6) "
	
	If Request("chkAddAll") <> "Y" Then
		sqlBase = sqlBase & "If @SaleType is null Begin " & _
							"set @SaleType = Case @UserType  " & _
							"		When 'C' Then (select ClientSaleUnit from olkcommon)  " & _
							"		When 'V' Then (select AgentSaleUnit from olkcommon) End End " & _
							"set @Quantity = " & getNumeric(Request("Quantity")) & " " & _
							"set @Quantity = (select @Quantity*Case @SaleType When 1 Then 1 When 2 Then NumInSale When 3 Then NumInSale*SalPackUn End from OITM where ItemCode = @ItemCode) "
	Else
		sqlBase = sqlBase & "set @Quantity = OLKCommon.dbo.DBOLKItemInvVal" & Session("ID") & "(@ItemCode, @WhsCode, @dbName, @LogNum, -1) "
	End If
	
		sqlBase = sqlBase & "declare @Unit int set @Unit = @SaleType " & _
				"declare @Price numeric(19,6) set @Price = " & getNumeric(Request.Form("precio")) & " "
	
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
	'response.redirect "http://www.topmanage.com.pa/query.asp?query=" & RetVal
End Function
%>

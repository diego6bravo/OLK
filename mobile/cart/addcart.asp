<% addLngPathStr = "cart/" %>
<!--#include file="lang/addcart.asp" -->
<!--#include file="../itemFunctions.asp"-->
<!--#include file="../lcidReturn.inc"-->
<%

If Request("cmd") = "itemdetails" Then

	If Request("Price") <> "no" then
		Price = "IsNull(Price, 0) Price, (IsNull(Price, 0) * numinsale) as salprice, itm1.Currency, "
		Price2 = "itm1.pricelist = " & Session("PList") & " and "
	End If
	sql = 	"select OITM.ItemCode, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'ItemName', OITM.ItemCode, OITM.ItemName) ItemName, " & _
			"Replace(Convert(nvarchar(4000),OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'UserText', OITM.ItemCode, UserText) collate database_default),Char(13),'<br>') Notes, " & Price & "PicturName, OITM.TreeType " & _
			"from oitm inner join itm1 on itm1.itemcode = oitm.itemcode " & _
			"where " & price2 & "SellItem = 'Y' and oitm.ItemCode = N'" & saveHTMLDecode(Request("Item"), False) & "'"
	set rs = conn.execute(sql)
	DocName = "" & getaddcartLngStr("LtxtDetails") & ""
Else
		If Session("PList") <> "" Then PriceList = Session("PList") Else PriceList = "NULL"
		If Session("RetVal") <> "" Then RetVal = Session("RetVal") Else RetVal = "NULL"
		
		sql = "declare @ItemCode nvarchar(20) set @ItemCode = N'" & saveHTMLDecode(Request("Item"), False) & "' " & _
		"declare @WhsCode nvarchar(8) set @WhsCode = OLKCommon.dbo.DBOLKGetWhsCode" & Session("ID") & "(" & Session("branch") & ", " & Session("vendid") & ", @ItemCode) " & _
		"declare @PriceList int set @PriceList = " & PriceList & " " & _
		"declare @LogNum Int set @LogNum = " & RetVal & " " & _
		"declare @LineNum int " & _
		"declare @LanID int set @LanID = " & Session("LanID") & " " & _
		"	set @LineNum = 	(select LineNum from R3_ObsCommon..DOC1 where ItemCode = @ItemCode and LogNum = @LogNum and LineNum = " & _
		"					(select Max(LineNum) from R3_ObsCommon..DOC1 where ItemCode = @ItemCode and LogNum = @LogNum)) " & _
		"declare @SaleType int  " & _
		"	set @SaleType = (select SaleType from OlkSalesLines where LogNum = @LogNum and LineNum = @LineNum) " & _
		"declare @DiscPrice numeric(19,6) declare @DiscExp nvarchar(10) declare @WithoutPriceList char(1) " & _
		"declare @UnitPrice numeric(19,6) " & _
		"EXEC OLKCommon..DBOLKGetItemDiscPrice" & Session("ID") & " N'" & saveHTMLDecode(Session("UserName"), False) & "', @ItemCode, @DiscPrice out, @DiscExp out, @WithoutPriceList out, null, @UnitPrice out " & _
		"select T0.ItemCode, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(@LanID, 'OITM', 'ItemName', T0.ItemCode, T0.ItemName) ItemName, " & _
					" OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(@LanID, 'OITM', 'SalUnitMsr', T0.ItemCode, T0.SalUnitMsr) SalUnitMsr, " & _
					" OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(@LanID, 'OITM', 'SalPackMsr', T0.ItemCode, T0.SalPackMsr) SalPackMsr, " & _
						"Replace(Convert(nvarchar(4000),OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(@LanID, 'OITM', 'UserText', T0.ItemCode, UserText)) collate database_default,Char(13),'<br>') Notes, PicturName, T0.TreeType, " & _
		"ISNULL(T2.Quantity,0) Verfy, " & _
		"@SaleType SaleType, NumInSale, SalPackUn, " & _
		"@WhsCode WhsCode, T2.Price CurPrice, ISNULL(IsNull(@DiscPrice,IsNull(@UnitPrice, T1.Price)),0) Price, ISNULL(IsNull(@DiscPrice,IsNull(@UnitPrice, T1.Price)),0)*NumInSale 'SalPrice', " & _
		"IsNull(Case @WithoutPriceList When 'N' Then IsNull(@UnitPrice, T1.Price) When 'Y' Then @DiscPrice End, IsNull(@UnitPrice, T1.Price)) UnitPrice, " & _
		"Case @WithoutPriceList When 'N' Then Case IsNull(@UnitPrice, T1.Price) When 0 Then 0 Else 100-(IsNull(@DiscPrice,IsNull(@UnitPrice, T1.Price))*100)/IsNull(@UnitPrice, T1.Price) End Else 0 End DiscPrcnt, " & _
		"Case IsNull(@UnitPrice, T1.Price) When 0 Then 0 Else 100-(IsNull(T2.Price,IsNull(@UnitPrice, T1.Price))*100)/IsNull(@UnitPrice, T1.Price) End CurDiscPrcnt, " & _
		"IsNull((select MaxDiscount from OLKAgentsAccess where SlpCode = " & Session("vendid") & "), 0) MaxLineDiscount " & _
		"from OITM T0 " & _
		"left outer join ITM1 T1 on T1.ItemCode = T0.ItemCode and T1.PriceList = @PriceList " & _
		"left outer join R3_ObsCommon..DOC1 T2 on T2.LogNum = @LogNum and T2.LineNum = @LineNum " & _
		"cross join OLKCommon " & _
		"where T0.ItemCode = @ItemCode "

set rs = conn.execute(sql)
  If Session("useraccess") = "U" Then
      MaxDiscount = rs("MaxLineDiscount")
  Else
  	MaxDiscount = 100
  End If

	If Request("err") = "disp" then
		vartext = "<font color=""#FF0000"" face=""Verdana"" style=""font-size: 7pt"">*" & getaddcartLngStr("LtxtNotEnoughQty") & "</font><hr noshade color=""#FF0000"" size=""1"">"
	ElseIf CDbl(Rs("Verfy")) = 0 Then
		vartext = ""
	ElseIf CDbl(Rs("Verfy")) > 0 Then ' and Not myApp.BasketMItems
		vartext = "<font color=""#FF0000"" face=""Verdana"" style=""font-size: 7pt"">*" & getaddcartLngStr("LtxtItmInCartNote") & "</font><hr noshade color=""#FF0000"" size=""1"">"
		CartItemFilter = True
	End If
	DocName = getaddcartLngStr("DtxtAdd")
	If Request("err") <> "disp" Then
		If Request("WhsCode") = "" Then WhsCode = rs("WhsCode") Else WhsCode = Request("WhsCode")
	Else 
		WhsCode = Request("WhsCode")
	End If
End If

If rs("PicturName") <> "" Then
	Pic = rs("PicturName")'
Else
	Pic = "n_a.gif"
End If 

TreeType = rs("TreeType")
		  
If Request("cmd") = "addcart" Then

If Request("Qty") <> "" Then 
	Qty = FormatNumber(CDbl(Request("Qty")), myApp.QtyDec) 
ElseIf Request("Quantity") <> "" Then
	Qty = Request("Quantity")
Else 
	Qty = FormatNumber(1, myApp.QtyDec) 
End If


If Request("SaleType2") = "" Then
	Unit = myApp.GetSaleUnit

	If CDbl(rs("NumInSale")) > 1 Then 
		If CDbl(rs("Verfy")) > 0 and rs("SaleType") = 1 or CDbl(rs("Verfy")) = 0 and myApp.GetSaleUnit = 1 Then
			Unit = 1
		End If
	End If
	If CDbl(rs("Verfy")) > 0 and rs("SaleType") = 2 or CDbl(rs("Verfy")) = 0 and myApp.GetSaleUnit = 2 or myApp.GetSaleUnit = 3 and CInt(Rs("SalPackUn")) = 1 Then 
		Unit = 2
	End If
	If CDbl(Rs("SalPackUn")) > 1 Then 
		If CDbl(rs("Verfy")) > 0 and rs("SaleType") = 3 or CDbl(rs("Verfy")) = 0 and myApp.GetSaleUnit = 3 Then
			Unit = 3
		End If
	End If
Else
	Unit = Request("SaleType2")
End If


If CDbl(rs("Verfy")) > 0 and Not myApp.BasketMItems Then
	Select Case rs("SaleType")
		Case 1
			SalePrice = rs("CurPrice")
			SaleType = 1
		Case 2
			SalePrice = rs("CurPrice")
			SaleType = 2
		Case 3
			If myApp.UnEmbPriceSet Then
				SalePrice = rs("CurPrice")
  			Else
				SalePrice = CDbl(rs("CurPrice"))*CDbl(rs("SalPackUn"))
  			End If
  			SaleType = 3
	End Select
  	DiscPrcnt = rs("CurDiscPrcnt")
Else
	Select Case CInt(myApp.GetSaleUnit)
		Case 1
			SalePrice = CDbl(RS("SalPrice")) / CDbl(rs("NumInSale"))
			If CDbl(rs("NumInSale")) > 1 Then SaleType = 1 Else SaleType = 2
		Case 2
			SalePrice = RS("SalPrice")
			SaleType = 2
		Case 3
			If myApp.UnEmbPriceSet Then
				SalePrice = RS("SalPrice")
			Else
				SalePrice = CDbl(RS("SalPrice"))*CDbl(rs("SalPackUn"))
			End If
			If CDbl(rs("SalPackUn")) > 1 Then SaleType = 3 Else SaleType = 2
	End Select
  	DiscPrcnt = rs("DiscPrcnt")
End If

If Request("SalePrice") <> "" Then 
	PriceVal = Replace(FormatNumber(CDbl(Request("SalePrice")), myApp.PriceDec), GetFormatSep, "") 
Else
	If Request("precio") = "" Then
		PriceVal = Replace(FormatNumber(SalePrice,myApp.PriceDec),GetFormatSep,"") 
	Else
		PriceVal = Request("precio")
		SaleType = Request("SaleType2")
	End If
End If

volSelBy = 1
If SaleType > 1 Then volSelBy = volSelBy * CDbl(rs("NumInSale"))
If SaleType = 3 Then volSelBy = volSelBy * CDbl(rs("SalPackUn")) %>
<script language="javascript">
var SaleType = <%=SaleType%>;
var NumInSale = <%=rs("NumInSale")%>;
var SalPackUn = <%=rs("SalPackUn")%>;
var varPrice = <%=GetNumeric(rs("Price"))%>;
var GetFormatDec = '<%=GetFormatDec()%>';
var UnEmbPriceSet = <%=JBool(myApp.UnEmbPriceSet)%>;
var SumDec = <%=myApp.SumDec%>;
var PercentDec = <%=myApp.PercentDec%>;
var PriceDec = <%=myApp.PriceDec%>;
var changePrice = true;
var txtValNumVal = '<%=getaddcartLngStr("DtxtValNumVal")%>';
var txtValNumMaxVal = '<%=getaddcartLngStr("DtxtValNumMaxVal")%>';
var txtMaxDiscount = '<%=getaddcartLngStr("LtxtMaxDiscount")%>';
var txtSelTaxCode = '<%=getaddcartLngStr("LtxtSelTaxCode")%>';
var MaxDiscount = <%=MaxDiscount%>;
function enableQty(chk)
{
	document.addcart.Quantity.disabled = chk;
}
function valAddFrm()
{
	<% If myApp.LawsSet = "MX" or myApp.LawsSet = "CL" or myApp.LawsSet = "CR" or myApp.LawsSet = "GT" or myApp.LawsSet = "US" or myApp.LawsSet = "CA" or myApp.LawsSet = "BR" Then
	TaxCode = getTaxCode
		If TaxCode <> "Disabled" Then %>
	    	if (document.addcart.TaxCode.selectedIndex == 0) {
	    		alert('<%=getaddcartLngStr("LtxtSelTaxCode")%>');
	    		document.addcart.TaxCode.focus();
	    		return false;
	    	}
	    <% End If
	End If %>
	document.addcart.action = 'cart/addcartsubmit.asp';
	return true;
}
<% 
set rd = Server.CreateObject("ADODB.RecordSet")
sql = "select T0.Amount,  " & _
"Case When AutoUpdt = 'N' Then T0.Price  " & _
"			Else T2.Price-(T2.Price*T0.Discount/100) " & _
"	End Price " & _
"from spp2 T0 " & _
"inner join spp1 T1 on T1.ItemCode = T0.ItemCode and T1.CardCode = T0.CardCode and T1.LineNum = T0.SPP1LNum " & _
"inner join ITM1 T2 on T2.ItemCode = T0.ItemCode and T2.PriceList = T1.ListNum " & _
"where (T0.CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "' or T0.CardCode = N'*"  & Session("plist") &"' and not exists(select 'A' " & _
"from spp2 T0 " & _
"inner join spp1 T1 on T1.ItemCode = T0.ItemCode and T1.CardCode = T0.CardCode and T1.LineNum = T0.SPP1LNum " & _
"inner join ITM1 T2 on T2.ItemCode = T0.ItemCode and T2.PriceList = T1.ListNum " & _
"where (T0.CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "') and T0.ItemCode = N'" & saveHTMLDecode(Request("Item"), False) & "' " & _
"and " & _
"  (T1.FromDate is null or DateDiff(day,getdate(),T1.FromDate) <= 0) and " & _
"  (T1.ToDate is null or DateDiff(day,getdate(),T1.ToDate) >= 0))) and T0.ItemCode = N'" & saveHTMLDecode(Request("Item"), False) & "' " & _
"and " & _
"  (T1.FromDate is null or DateDiff(day,getdate(),T1.FromDate) <= 0) and " & _
"  (T1.ToDate is null or DateDiff(day,getdate(),T1.ToDate) >= 0) "
set rd = conn.execute(sql)
If rd.Eof Then %>
var itemVolDisc = null;
<% Else
do while not rd.eof
	If itemVolDisc <> "" Then itemVolDisc = itemVolDisc & "{S}"
	itemVolDisc = itemVolDisc & rd("Amount") & "|" & CDbl(rd("Price"))
rd.movenext
loop %>
var itemVolDisc = '<%=itemVolDisc%>';
<% End If %>
</script>
<script type="text/javascript" src="cart/addcart.js"></script>
<% End If

      set rx = Server.CreateObject("ADODB.Recordset")
      set rxVal = Server.CreateObject("ADODB.Recordset")
      sql = "select T0.rowIndex, IsNull(T1.alterRowName, T0.rowName) rowName, T0.rowField, T0.rowType, T0.rowTypeRnd, T0.rowTypeDec, T0.HideNull, T0.linkActive, T0.linkObject,  " & _
      	"T2.rsName, Case When T2.rsIndex is not null Then 'Y' Else 'N' End Verfy " & _
		"from olkitemrep T0 " & _
		"left outer join OLKItemRepAlterNames T1 on T1.rowIndex = T0.rowIndex and T1.LanID = " & Session("LanID") & " " & _
      	"left outer join OLKRS T2 on T2.rsIndex = T0.linkObject " & _
		"where T0.rowAccess in ('T','V') and T0.rowOP in ('T','P') "
		
      If Session("username") = "" Then sql = sql & " and Convert(varchar(8000),rowField) not like '%@CardCode%' "
      
      If Request("cmd") <> "addcart" Then sql = sql & " and Convert(varchar(8000),rowField) not like '%@Quantity%' and Convert(varchar(8000),rowField) not like '%@Price%' and Convert(varchar(8000),rowField) not like '%@Unit%' "
      
      If TreeType <> "S" Then
      	sql = sql & " and T0.rowIndex <> -1 "
      End If
            
      sql = sql & " order by rowOrder "
      rx.open sql, conn, 3, 1   
      If Rx.RecordCount > 0 Then
      If Session("plist") <> "" Then PriceList = " declare @PriceList int set @PriceList = " & Session("plist")
      If Session("UserName") <> "" Then CardCode = " declare @CardCode nvarchar(20) set @CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "'"
      		sqlx = ""
      		 do while not rx.eof
      		 	varx = varx + 1
      		 	rowName = Replace(Rx("rowName"), "'", "''")
      		 	
      		 	If PriceList = "" and InStr(rx("rowField"),"Price") <> 0 Then
      		 	ElseIf CardCode = "" and InStr(rx("rowField"),"CardCode") <> 0 Then
      		 	Else
      		 		If sqlx <> "" Then sqlx = sqlx & ", "
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
	      		 	Select Case rx("rowType") 
	      		 		Case "L" 
			      		 	sqlx = sqlx & " OLKCommon.dbo.DBOLKCode" & Session("ID") & "('L'," & rowTypeRnd & "Convert(nvarchar(20),(" & Rx("rowField") & ")), " & myDec & ")" & " As N'" & rowName & "'"
			      		 Case "M" 
			      		 	sqlx = sqlx & " OLKCommon.dbo.DBOLKCode" & Session("ID") & "('M'," & rowTypeRnd & "Convert(nvarchar(20),(" & Rx("rowField") & ")), " & myDec & ")" & " As N'" & rowName & "'"
			      		 Case "H" 
			      		 	sqlx = sqlx & " OLKCommon.dbo.DBOLKCode" & Session("ID") & "('H'," & rowTypeRnd & "Convert(nvarchar(20),(" & Rx("rowField") & ")), " & myDec & ")" & " As N'" & rowName & "'"
			      		 Case "F" 
			      		 	sqlx = sqlx & Rx("rowField") & " As N'" & rowName & "'"
		      		 	 Case Else
			      		 	sqlx = sqlx & "(" & Rx("rowField") & ") As N'" & rowName & "'"
	      		 	End Select
	      		 End If
      		 rx.movenext
      		 loop
      		sql = PriceList & CardCode & _
      			   " declare @SlpCode int set @SlpCode = " & Session("vendid") & _
      			   " declare @dbName nvarchar(100) set @dbName = '" & Session("OlkDB") & "'" & _
      			   " declare @ItemCode nvarchar(20) set @ItemCode = N'" & saveHTMLDecode(Request("item"), False) & "' " & _
      			   " declare @LanID int set @LanID = " & Session("LanID") & " " & _
				   " declare @WhsCode nvarchar(8) set @WhsCode = OLKCommon.dbo.DBOLKGetWhsCode" & Session("ID") & "(" & Session("branch") & ", " & Session("vendid") & ", @ItemCode) " & _
				   " declare @branchIndex int set @branchIndex = " & Session("branch") & " "
				   
			If Request("cmd") = "addcart" Then
				sql = sql & "declare @Quantity numeric(19,6) set @Quantity = " & getNumeric(Qty) & " " & _
							"declare @Price numeric(19,6) set @Price = " & getNumeric(PriceVal) & " " & _
							"declare @Unit int set @Unit = " & Unit & " "
			End If
				   
      		sql = sql & " select " & sqlx & " from oitm where itemcode = N'" & saveHTMLDecode(Request("item"), False) & "'"
      		sql = QueryFunctions(sql)
      		rxVal.open sql, conn, 3, 1
      End If
 %>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
    <tr>
      <td bgcolor="#9BC4FF">
      <form method="POST" action="cart/addcartsubmit.asp" name="addcart">
        <table border="0" cellpadding="0" bordercolor="#111111" width="100%">
        <tr>
          <td><font face="Verdana" size="2"><%=getaddcartLngStr("DtxtItem")%>:&nbsp;<string><%=Request("Item")%></string></font></td>
        </tr>
        <tr>
          <td>
          <table border="0" cellpadding="0" cellspacing="1"  bordercolor="#111111" width="100%">
            <tr><% If myApp.ShowPocketImg Then %>
              <td width="86" valign="top">
                <a href="operaciones.asp?cmd=viewImage&amp;FileName=<%=Pic%>"><img border="0" src="pic.aspx?filename=<%=Pic%>&dbName=<%=Session("olkdb")%>&MaxSize=80"></a></td><% End If %>
              <td valign="top"<% If myApp.ShowPocketImg Then %> width="154"<% End If %>><font size="1" face="Verdana"><% If CartItemFilter Then %><a href="#" onclick="javascript:goCartFilter();"><img border="0" src="images/<%=Session("rtl")%>flechaselec.gif" align="<% If Session("rtl") = "" Then %>left<% Else %>right<% End If %>"></a><% End If %><%=vartext%><%=rs("Notes") %></font></td>
            </tr>
          </table>
          </td>
        </tr>
		<tr>
			<td>
				<table cellpadding="0" cellspacing="0" border="0" width="100%">
					<tr>
						<td valign="top">
						<% If Request("cmd") = "addcart" Then %>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td>
								<font size="1" face="Verdana"><b><%=getaddcartLngStr("DtxtWarehouse")%></b></font>
								</td>
								<td colspan="2">
								<% 
								set cmd = Server.CreateObject("ADODB.Command")
								cmd.ActiveConnection = connCommon
								cmd.CommandType = &H0004
								cmd.CommandText = "DBOLKGetWarehouses" & Session("ID")
								cmd.Parameters.Refresh()
								cmd("@LanID") = Session("LanID")
								If myAut.HasAuthorization(99) Then %>
								<select name="WhsCode" size="1" class="input" style="font-size: 10px; font-family:Verdana">
								<% 
								set rw = cmd.execute()
								do while not rw.eof %>
								<option <% If CStr(WhsCode) = CStr(rw("WhsCode")) then response.write "selected " %>value="<%=myHTMLEncode(rw("WhsCode"))%>"><%=myHTMLEncode(rw("WhsName"))%></option>
								<% rw.movenext
								loop %>
								</select>
								<% Else
								cmd("@Filter") = WhsCode
								set rw = cmd.execute() %>
								<input type="hidden" name="WhsCode" value="<%=WhsCode%>">
								<font size="1" face="Verdana"><%=rw("WhsName")%></font>
								<% End If %>
								</td>
							</tr>
							<tr>
								<td>
								<input type="hidden" name="SaleType" value="<%=SaleType%>">
								<INPUT TYPE="HIDDEN" NAME="Item" VALUE="<%=Request("Item")%>">
								<INPUT TYPE="HIDDEN" NAME="PackPrice" VALUE="<%=Request("PackPrice")%>">
								<font size="1" face="Verdana"><b><%=getaddcartLngStr("DtxtQty")%></b></font>
								</td>
								<td><% If myApp.EnableUnitSelection Then %>
								<font face="Verdana" size="1"><b><%=getaddcartLngStr("DtxtSalUnit")%></b></font>
								<% If CDbl(rs("Verfy")) > 0 and Not myApp.BasketMItems Then %><input type="hidden" name="SaleType2" value="<%=rs("SaleType")%>">
								<input name="precio" type="hidden" value="<% If Request("SalePrice") <> "" Then Response.write Replace(FormatNumber(CDbl(Request("SalePrice")), myApp.PriceDec), ",", "") Else Response.write Replace(FormatNumber(SalePrice,myApp.PriceDec),",","") %>" ><% End If %>
								<% Else %>&nbsp;<input type="hidden" name="SaleType2" value="<%=myApp.GetSaleUnit%>"><% End If %></td>
								<% If myApp.ShowLineDiscount Then %>
								<td><font face="Verdana" size="1"><b><%=getaddcartLngStr("DtxtDiscount")%></b></font></td>
								<% End If %>
							</tr>
            				<tr>
            					<td>
            					<span dir="rtl">
            					<input class="input" name="Quantity" type="number" min="0" step="<%=GetNumberStep(myApp.QtyDec)%>" size="6" style="font-size: 10px; font-family:Verdana; text-align:right" value="<%=Qty%>" onclick="this.selectionStart=0;this.selectionEnd=this.value.length;" onchange="javascript:chkThis(this, document.addcart.oldQuantity, <%=myApp.QtyDec%>, 32999999999.00)">
            					<input type="hidden" name="oldQuantity" value="<%=Qty%>">
            					</span>
            					</td>
								<td>
								<% If myApp.EnableUnitSelection Then %>
								<select <% If CDbl(rs("Verfy")) > 0 and Not myApp.BasketMItems Then %>disabled<%end if%> class="input" size="1" name="SaleType<% If CDbl(rs("Verfy")) > 0 and Not myApp.BasketMItems Then %>x<%end if%>2" onchange="javascript:changeUnEmb(this, document.addcart.precio);<% If Request("ShowVolDisc") = "Y" Then %>document.addcart.action = 'operaciones.asp';submit();<% End If %>" style="font-size: 10px; font-family:Verdana">
								<option value="1" <% If Unit = 1 Then %>selected<%end if%>><%=getaddcartLngStr("DtxtUnit")%>(1)</option>
								<option value="2" <% If Unit = 2 Then %>selected<%end if%>><%=rs("SalUnitMsr")%><% If myApp.GetShowQtyInUn Then %>(<%=rs("NumInSale")%>)<% End If %></option>
								<option value="3" <% If Unit = 3 Then %>selected<%end if%>><%=rs("SalPackMsr")%><% If myApp.GetShowQtyInUn Then %>(<%=rs("SalPackUn")%>)<% End If %></option>
								</select><% Else %>&nbsp;<input type="hidden" name="SaleType" value="<%=myApp.GetSaleUnit%>"><% End If %>
								</td>
								<% If myApp.ShowLineDiscount Then %>
								<td><input <% If Not myAut.HasAuthorization(68) Then %>readonly <% End If %> <% If rs("CurPrice") <> "" and Not myApp.BasketMItems Then %> disabled<%end if%> name="DiscPrcnt" size="10" onclick="this.selectionStart=0;this.selectionEnd=this.value.length;" value="<%=FormatNumber(DiscPrcnt, myApp.PercentDec)%>" onchange="javascript:chkDiscount(this);" style="font-size: 10px; font-family:Verdana; text-align:right" type="number" step="<%=GetNumberStep(myApp.PercentDec)%>">
				            	</td>
				            	<% Else %>
				             	<input type="hidden" name="DiscPrcnt" value="<%=FormatNumber(DiscPrcnt, myApp.PercentDec)%>">
				            	<% End If %>
				            	<input type="hidden" name="oldDiscPrcnt" value="<%=DiscPrcnt%>">
							</tr>
							<% If myApp.EnSelAll Then %>
							<tr>
							<td colspan="3"><input type="checkbox" name="chkAddAll" value="Y" id="chkAddAll" style="border-style:solid; border-width:0px;" onclick="enableQty(this.checked);"><label for="chkAddAll"><font size="1" face="Verdana"><%=getaddcartLngStr("LtxtSellAll")%></font></label></td>
							</tr>
							<% End If %>
							<% If itemVolDisc <> "" and Request("ShowVolDisc") = "Y" Then %>
							<tr>
								<td colspan="3">
								<table border="0" width="100%">
									<tr>
										<td colspan="2" bgcolor="#75ACFF" align="center"><b><font size="1" face="Verdana"><%=getaddcartLngStr("LtxtVolDiscount")%></font></b></td>
									</tr>
									<tr>
										<td width="50%" bgcolor="#75ACFF"><b><font size="1" face="Verdana"><%=getaddcartLngStr("DtxtQty")%></font></b></td>
										<td width="50%" bgcolor="#75ACFF"><b><font size="1" face="Verdana"><%=getaddcartLngStr("DtxtPrice")%></font></b></td>
									</tr>
									<% rd.movefirst
									do while not rd.eof %>
									<tr>
										<td width="50%" align="right" bgcolor="#75ACFF"><font size="1" face="Verdana"><%=rd("Amount")%></font></td>
										<td width="50%" align="right" bgcolor="#75ACFF"><font size="1" face="Verdana"><%=FormatNumber(CDbl(rd("Price"))*volSelBy, myApp.PriceDec)%></font></td>
									</tr>
									<% rd.movenext
									loop %>
								</table>
								</td>
							</tr>
							<% End If %>
				            <tr>
				            	<td>
				            	<table cellpadding="0" cellspacing="0" border="0">
				            		<tr>
				            			<td><font face="Verdana" size="1"><b><%=getaddcartLngStr("DtxtPrice")%></b></font></td>
				            			<% If itemVolDisc <> "" Then %><td align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
										<input type="image" src="../images/foco_icon.gif" width="23" height="22" style="vertical-align: middle" onclick="javascript:document.addcart.action = 'operaciones.asp';"></td><% End If %>
				            		</tr>
				            	</table>
								</td>
								<td>
								<font face="Verdana" size="1"><b><%=getaddcartLngStr("DtxtTotal")%></b></font>
								&nbsp;</td>
								<td>
								<% If myApp.LawsSet = "MX" or myApp.LawsSet = "CL" or myApp.LawsSet = "CR" or myApp.LawsSet = "GT" or myApp.LawsSet = "US" or myApp.LawsSet = "CA" or myApp.LawsSet = "BR" Then %><font face="Verdana" size="1"><b><%=getaddcartLngStr("LtxtTaxCode")%></b></font><% End If %></td>
								</tr> 
				            	<tr>
				            		<td>
				            		<font face="Verdana" size="1"><b><span dir="rtl">
				            		<input type="number" min="0" step="<%=GetNumberStep(myApp.PriceDec)%>" <% If rs("CurPrice") <> "" and Not myApp.BasketMItems Then %>disabled<%end if%>  <% If Not myAut.HasAuthorization(68) Then %> readonly <% End If %> name="precio<% If rs("CurPrice") <> "" and Not myApp.BasketMItems Then %>x<%end if%>" size="10" onclick="this.selectionStart=0;this.selectionEnd=this.value.length;" value="<%=PriceVal%>"  onchange="javascript:chkThis(this, document.addcart.oldPrecio, <%=myApp.PriceDec%>, 8699999999999.000)" style="font-size: 10px; font-family:Verdana; text-align:right; height: 18px;">
				            		<input type="hidden" name="oldPrecio" value="<%=PriceVal%>">
				            		</span></b></font>
				            		</td>
									<td>
									<font face="Verdana"><b><span dir="rtl">
									<input readonly <% If rs("CurPrice") <> "" Then %>disabled<%end if%> name="LineTotal" size="10" style="font-size: 10px; font-family:Verdana; text-align:right">
									</span></b></font></td>
									<td>
									<% If (myApp.LawsSet = "MX" or myApp.LawsSet = "CL" or myApp.LawsSet = "CR" or myApp.LawsSet = "GT" or myApp.LawsSet = "US" or myApp.LawsSet = "CA" or myApp.LawsSet = "BR") Then
									If myAut.HasAuthorization(175) Then %>
									<select class="input" <% If rs("CurPrice") <> "" and Not myApp.BasketMItems Then %>disabled<%end if%> name="TaxCode">
									<% If TaxCode <> "Disabled" Then %>
									<option></option>
									<% 
									set rw = Server.CreateObject("ADODB.RecordSet")
									sql = "select Code from OSTC where ValidForAR = 'Y'"
									set rw = conn.execute(sql)
									do while not rw.eof %>
									<option <% If TaxCode = rw("Code") Then %>selected<% End If %> value="<%=rw("Code")%>"><%=rw("Code")%></option>
									<% rw.movenext
									loop
									Else %>
									<option value="">|L:txtNotApply|</option>
									<% End If %>
									</select>
									<% Else %>
									<%=TaxCode%><input type="hidden" name="TaxCode" value="<%=TaxCode%>">
									<% End If %>
									<% Else %>&nbsp;<% End If %></td>
								</tr>
								<tr>
									<td colspan="3">
									<table border="0" cellpadding="0" width="100%" id="table2">
										<tr>
											<% If Request("cmd") = "addcart" Then %>
											<td>
											<input type="image" name="I1" border="0" src="images/ok_icon.gif" onclick="javascript:return valAddFrm();"></td>
											<% End If %>
											<td align="center"><a href='operaciones.asp?cmd=slistsearch<% If Request("cmd") = "itemdetails" Then %>&amp;slist=Y<% End If %>'><img name="pocket_art_r3_c3" src="images/search_icon.gif" border="0"></a></td>
											<% If Request("cmd") = "addcart" Then %>
											<td align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><a href="javascript:javascript:if(confirm('<%=getaddcartLngStr("LtxtConfCancel")%>'))<% If Request("retSearch") <> "Y" Then %>history.go(-1);<% Else %>window.location.href='operaciones.asp?cmd=slistsearch';<% End If %>"><img border="0" src="images/x_icon.gif"></a></td>
											<% End If %>
										</tr>
									</table>
									</td>
								</tr>
							</table>	
							<% Else %>
							<a href='operaciones.asp?cmd=slistsearch<% If Request("cmd") = "itemdetails" Then %>&amp;slist=Y<% End If %>'><img name="pocket_art_r3_c3" src="images/search_icon.gif" border="0"></a>
							<% End If %>	
						</td>
					</tr>
				</table>
			</td>
		</tr>
        <tr>
          <td>
          <table border="0" cellpadding="0"  bordercolor="#111111" width="100%">
            <tr>
              <td bgcolor="#75ACFF" align="left" valign="top"><b><font size="1" face="Verdana"><%=getaddcartLngStr("DtxtDescription")%></font></b></td>
              <td><font face="Verdana" size="1"><% =rs("ItemName") %>&nbsp;</font></td>
            </tr>
            <% If Request("cmd") = "itemdetails" and Request("Price") <> "no" Then %>
            <tr>
              <td bgcolor="#75ACFF" align="left" valign="top"><b><font size="1" face="Verdana"><%=getaddcartLngStr("DtxtPrice")%></font></b></td>
              <td><font face="Verdana" size="1"><nobr><%=rs("Currency")%>&nbsp;<% =FormatNumber(CDbl(rs("salprice")), myApp.PriceDec) %></nobr></font></td>
            </tr>
            <% End If %>
    		<% If rx.recordcount > 0 Then
    		For Each Field in rxVal.Fields
    		rx.Filter = "rowName = '" & Field.Name & "'"
    		customVal = rxVal(Field.name)
    		If Not IsNull(customVal) or IsNull(customVal) and rx("HideNull") = "N" Then %>
            <tr>
              <td bgcolor="#75ACFF" align="left" valign="top">
		        <table cellpadding="0" cellspacing="0" border="0" width="100%">
		        	<tr>
		        		<td><font size="1" face="Verdana"><b><%=Field.name%></b></font></td>
		        		<% If rx("linkActive") = "Y" Then %><td width="15">
		        		<a href="javascript:<% If rx("Verfy") = "Y" Then %>doLink(<%=rx("rowIndex")%>)<% Else %>doErrRep()<% End If %>;">
						<img alt="<%=myHTMLEncode(rx("rsName"))%>" border="0" src="images/<%=Session("rtl")%>flechaselec.gif" width="15" height="13" style="cursor: hand"></a></td><% End If %>
		        	</tr>
		        </table>
              </td>
              <td><font size="1" face="Verdana"><%=customVal%></font></td>
            </tr>
  			<% 
  			End If
  			Next
  	 		end if %>
          </table>
          </td>
        </tr>
        <% If myAut.HasAuthorization(100) or myAut.HasAuthorization(102) or myAut.HasAuthorization(103) Then %>
        <tr>
          <td>
          <table border="0" cellpadding="0" cellspacing="0" width="100%" id="table1">
			<tr>
				<td align="center">
				<p align="left">&nbsp;<a href="#" onclick="javascript:window.location.href='operaciones.asp?cmd='+document.addcart.report.value+'&item=<%=Request("Item")%>'"><img border="0" src="images/<%=Session("rtl")%>flechaselec.gif"></a><select size="1" name="report" style="font-family: Verdana; font-size: 10px">
				<% If myAut.HasAuthorization(103) Then %><option value="olkRep"><%=getaddcartLngStr("LtxtCommItems")%></option><% End If %>
				<% If myAut.HasAuthorization(100) Then %><option value="bpriceRep"><%=getaddcartLngStr("LtxtBestPrice")%></option><% End If %>
				<% If myAut.HasAuthorization(102) Then %><option value="ventasRep"><%=getaddcartLngStr("Ltxt10LastSales")%></option><% End If %>
				</select></td>
			</tr>
			</table>
			</td>
        </tr>
        <% End If %>
        </table>
        <input type="hidden" name="PackPrice" value="<%=Request("PackPrice")%>">
        <input type="hidden" name="ShowVolDisc" value="Y">
        <input type="hidden" name="cmd" value="addcart">
      	<input type="hidden" name="ManPrc" value="<% If Request("ManPrc") = "Y" Then %>Y<% Else %>N<% End If %>">
      	<% If Request("cmd") <> "itemdetails" Then %><input type="hidden" name="UnitPrice" value="<%=rs("UnitPrice")%>"><% End If %>
      </form>
      </td>
    </tr>
    </table>
  </center>
</div>
<form name="frmGoCartFilter" method="post" action="operaciones.asp">
<input type="hidden" name="string" value="<%=Request("Item")%>">
<input type="hidden" name="cmd" value="cart">
<input type="hidden" name="document" value="B">
</form>
<%
set rx = Server.CreateObject("ADODB.RecordSet")
set rxVal = Server.CreateObject("ADODB.RecordSet")

sql = "select T0.rowIndex, T1.rsIndex " & _  
"from OLKItemRep T0 " & _  
"inner join OLKRS T1 on T1.rsIndex = T0.linkObject " & _  
"where T0.rowAccess in ('T','V') and T0.rowOP in ('T','P') and linkActive = 'Y' " 

rx.open sql, conn, 3, 1

sql = "select T0.rowIndex, T1.varIndex, T1.varVar, T1.varDataType, T2.valBy, T2.valValue, T2.valDate " & _  
"from OLKItemRep T0 " & _  
"inner join OLKRSVars T1 on T1.rsIndex = T0.linkObject " & _  
"left outer join OLKItemRepLinksVars T2 on T2.rowIndex = T0.rowIndex and T2.varId = T1.varVar " & _  
"where T0.rowAccess in ('T','V') and T0.rowOP in ('T','P') and linkActive = 'Y' " 
rxVal.open sql, conn, 3, 1

%>
<script language="javascript">
<% If Not rx.eof Then %>
function doLink(rowIndex)
{
	switch (rowIndex)
	{
		<% do while not rx.eof %>
		case <%=rx("rowIndex")%>:
			document.frmLink<%=Replace(rx("rowIndex"), "-", "_")%>.submit();
			break;
		<% rx.movenext
		loop %>
	}
}
function doErrRep()
{
	alert('<%=getaddcartLngStr("LtxtErrRep")%>');
}
<% rx.movefirst
End If %>
function goCartFilter()
{
	document.frmGoCartFilter.submit();
}
<% If Request("cmd") = "addcart" Then %>
setLineTotal()
<% End If %>
</script>
<% do while not rx.eof %>
<form name="frmLink<%=Replace(rx("rowIndex"), "-", "_")%>" id="frmLink<%=Replace(rx("rowIndex"), "-", "_")%>" action="operaciones.asp" method="post">
<input type="hidden" name="rsIndex" value="<%=rx("rsIndex")%>">
<input type="hidden" name="itemCmd" value="A">
<input type="hidden" name="cmd" value="viewRep">
<%
rxVal.Filter = "rowIndex = " & rx("rowIndex")
do while not rxVal.eof
Select Case rxVal("valBy") 
	Case "V"
		If rxVal("varDataType") <> "datetime" Then
			strVal = rxVal("valValue")
		Else
			strVal = rxVal("valDate")
		End If
	Case "F"
		Select Case rxVal("valValue")
			Case "@PriceList"
				strVal = Session("plist")
			Case "@SlpCode"
				strVal = Session("vendid")
			Case "@CardCode"
				strVal = saveHTMLDecode(Session("UserName"), False)
			Case "@WhsCode"
				strVal = saveHTMLDecode(WhsCode, False)
			Case "@dbName"
				strVal = Session("olkdb")
			Case "@ItemCode"
				strVal = saveHTMLDecode(Request("item"), False)
			Case "@Quantity"
				strVal = getNumeric(Qty)
			Case "@Unit"
				strVal = Unit
			Case "@Price"
				strVal = getNumeric(PriceVal)
		End Select
	Case "Q"
		If Session("plist") <> "" Then PriceList = " declare @PriceList int set @PriceList = " & Session("plist")
      	sql = PriceList & _
      			   " declare @SlpCode int set @SlpCode = " & Session("vendid") & _
      			   " declare @CardCode nvarchar(20) set @CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "'" & _
      			   " declare @WhsCode nvarchar(8) set @WhsCode = N'" & saveHTMLDecode(WhsCode, False) & "'" & _
      			   " declare @dbName nvarchar(100) set @dbName = N'" & Session("OlkDB") & "' " & _
      			   " declare @ItemCode nvarchar(20) set @ItemCode = '" & saveHTMLDecode(Request("item"), False) & "' " & _
      			   " select (" & rxVal("valValue") & ") "
		set rv = conn.execute(sql)
		If Not rv.Eof Then strVal = rv(0) Else strVal = ""
End Select
 %>
<input type="hidden" name="var<%=rxVal("varIndex")%>" value="<%=myHTMLEncode(strVal)%>">
<% rxVal.movenext
loop %>
</form>
<% rx.movenext
loop
set rxVal = nothing
set rx = nothing %>
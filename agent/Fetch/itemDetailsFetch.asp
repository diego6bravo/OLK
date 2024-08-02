<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%
Response.Expires = -1
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="../chkLogin.asp" -->
<!--#include file="../myHTMLEncode.asp"-->
<!--#include file="../lcidReturn.inc" -->
<!--#include file="../authorizationClass.asp"-->
<%

Dim myAut
set myAut = New clsAuthorization

DataType = Request("DataType")
ItemCode = Request("Item")

VirtualTotal = 0
VirtualChkInv = "Y"

set rs = Server.CreateObject("ADODB.RecordSet")
Select Case DataType 
	Case "D" 'Item Data
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKItemDetailsFetchData" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@ItemCode") = ItemCode
		cmd("@branchIndex") = Session("branch")
		cmd("@SlpCode") = Session("vendid")
		cmd("@SaleType") = myApp.GetSaleUnit
		cmd("@BasketMItems") = GetYN(myApp.BasketMItems)
		cmd("@UnEmbPriceSet") = GetYN(myApp.UnEmbPriceSet)
		cmd("@LawsSet") = myApp.LawsSet
		
		cmd("@cmd") = Request("cmd")
		If Request("cmd") = "A" Then 
			cmd("@LogNum") = Session("RetVal")
			cmd("@CardCode") = Session("UserName")
			cmd("@PriceList") = Session("PriceList")
		End If
		
		If Request("LineNum") <> "" Then cmd("@LineNum") = CInt(Request("LineNum"))
		
		set rs = cmd.execute()

		If rs("SaleType") = 3 Then FormatQty = 0 Else FormatQty = myApp.QtyDec
		
		For i = 0 to rs.Fields.Count - 1
			Select Case rs.Fields(i).Name
				Case "Quantity"
					If Not myApp.BasketMItems and CDbl(rs("VerfyOnCart")) > 0 Then 
						If Not IsNull(rs("Quantity")) Then Response.Write FormatNumber(CDbl(rs("Quantity")), FormatQty)
					End If
				Case "Price"
					If Not IsNull(rs("Price")) Then Response.Write FormatNumber(CDbl(rs("Price")), myApp.PriceDec)
				Case "DiscPrcnt"
					DiscPrcnt = 0
					If Not IsNull(rs("DiscPrcnt")) Then 
						DiscPrcnt = CDbl(rs("DiscPrcnt"))
						Response.Write FormatNumber(DiscPrcnt, myApp.PercentDec)
					End If
				Case "VerfyOnCart"
					Response.Write GetYN(CDbl(rs("VerfyOnCart")) > 0)
				Case "Currency"
					If Not IsNull(rs("Currency")) Then Response.Write rs("Currency") Else Response.Write myApp.MainCur
				Case Else
					Response.Write rs(i)
			End Select
			Response.Write "{S}"
		Next
		
		If Not IsNull(rs("VerfyOnCart")) Then Response.Write FormatNumber(CDbl(rs("VerfyOnCart")), myApp.QtyDec)
		Response.Write "{S}"
		CheckDiscount rs("TreeType"), rs("WhsCode"), 1, myApp.GetSaleUnit, rs("OLKCombo"), rs("ShowComp"), rs("Virtual"), DiscPrcnt
		Response.Write "{S}"
		If rs("Virtual") = "Y" Then Response.Write FormatNumber(VirtualTotal, myApp.PriceDec)
		Response.Write "{S}"
		If rs("Virtual") = "Y" Then Response.Write FormatNumber(VirtualTotal, myApp.SumDec)
		Response.Write "{S}"
		If rs("Virtual") = "Y" Then Response.Write VirtualChkInv
		Response.Write "{S}"
		If rs("Virtual") = "Y" Then Response.Write FormatNumber(0, myApp.PercentDec)
		Response.Write "{S}"
		Response.Write FormatNumber(1, FormatQty)
		Response.Write "{S}"
		If Not myApp.BasketMItems and CDbl(rs("VerfyOnCart")) > 0 Then Response.Write "Y" Else Response.Write "N"
		Response.Write "{S}"
		If Not myApp.UnEmbPriceSet or Not myApp.UnEmbPriceSet and myApp.GetSaleUnit < 3 Then
			If Not IsNull(rs("Price")) Then Response.Write FormatNumber(CDbl(rs("Price")), myApp.SumDec)
		Else
			If Not IsNull(rs("Price")) Then Response.Write FormatNumber(CDbl(rs("Price"))*CDbl(rs("SalPackUn")), myApp.SumDec)
		End If
	Case "WR"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKSalesItemDetails" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@ItemCode") = ItemCode
		set rs = cmd.execute()
		
		With Response
			.Write FormatNumber(RS("OnHand"),0)
			.Write "{C}"
			.Write FormatNumber(RS("DispSAP"),0)
			.Write "{C}"
			.Write FormatNumber(RS("InvOLKDisp"),0)
			.Write "{C}"
			.Write FormatNumber(RS("OnHandUnVentSAP"),0) & " (" & FormatNumber(RS("OnHandSueltoUnVentaSAP"),0) & ")"
			.Write "{C}"
			.Write FormatNumber(RS("DispUnVentSAP"),0) & " (" & FormatNumber(RS("DispSueltoUnVentSAP"),0) & ")"
			.Write "{C}"
			.Write FormatNumber(RS("InvOLKUnVentDisp"),0) & " (" & FormatNumber(RS("InvOLKSueltoUnVentDisp"),0) & ")"
			.Write "{C}"
			.Write FormatNumber(RS("OnHandUnEmbSAP"),0) & " (" & FormatNumber(RS("OnHandSueltoUnEmbSAP"),0) & ")"
			.Write "{C}"
			.Write FormatNumber(RS("DispUnEmbSAP"),0) & " (" & FormatNumber(RS("DispSueltoUnEmbSAP"),0) & ")"
			.Write "{C}"
			.Write FormatNumber(RS("InvOLKUnEmbDisp"),0) & " (" & FormatNumber(RS("InvOLKSueltoUnEmbDisp"),0) & ")"
			
			.Write "{S}"
			
			arrWhs = Split(Request("WhsCode"), ", ")
			For i = 0 to UBound(arrWhs)
				cmd("@WhsCode") = arrWhs(i)
				set rs = cmd.execute()
				
				If i > 0 Then .Write "{W}"
				
				.Write FormatNumber(RS("InvBDGWhs"),0)
				.Write "{C}"
				.Write FormatNumber(RS("InvBDGDisp"),0)
				.Write "{C}"
				.Write FormatNumber(RS("InvOLKBDGDisp"),0)
				.Write "{C}"
				.Write FormatNumber(RS("InvUnVentBDGWhs"),0) & " (" & FormatNumber(RS("InvSueltoUnVentBDGWhs"),0) & ")"
				.Write "{C}"
				.Write FormatNumber(RS("InvBDGUnVentDisp"),0) & " (" & FormatNumber(RS("InvBDGSueltoUnVentDisp"),0) & ")"
				.Write "{C}"
				.Write FormatNumber(RS("InvOLKBDGUnVentDisp"),0) & " (" & FormatNumber(RS("InvOLKBDGSueltoUnVentDisp"),0) & ")"
				.Write "{C}"
				.Write FormatNumber(RS("InvUnEmbBDGWhs"),0) & " (" & FormatNumber(RS("InvSueltoUnEmbBDGWhs"),0) & ")"
				.Write "{C}"
				.Write FormatNumber(RS("InvBDGUnEmbDisp"),0) & " (" & FormatNumber(RS("InvBDGSueltoUnEmbDisp"),0) & ")"
				.Write "{C}"
				.Write FormatNumber(RS("InvOLKBDGUnEmbDisp"),0) & " (" & FormatNumber(RS("InvOLKBDGSueltoUnEmbDisp"),0) & ")"
			Next
		End With
		
	Case "SR" 'Sales Report
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKItemDetailsFetchSalesRep" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@ItemCode") = ItemCode
		If Not myAut.HasAuthorization(97) Then cmd("@SlpCode") = Session("vendid")
		cmd("@PriceDec") = myApp.PriceDec
		set rs = cmd.execute()
		
		If Not rs.Eof Then
			WriteSep = False
			do while not rs.eof
				With Response
					If WriteSep Then Response.Write "{S}"
					.Write rs("DocEntry")
					.Write "{C}"
					.Write rs("LineNum")
					.Write "{C}"
					.Write rs("DocNum") & " (" & (CInt(rs("LineNum"))+1) & ")"
					.Write "{C}"
					.Write FormatDate(rs("DocDate"), true)
					.Write "{C}"
					.Write rs("CardCode")
					.Write "{C}"
					.Write rs("CardName")
					.Write "{C}"
					.Write FormatNumber(CDbl(rs("Quantity")), myApp.QtyDec)
					.Write "{C}"
					.Write rs("MType")
					If rs("UseBaseUn") = "N" and myApp.GetShowQtyInUn Then .Write "(" & rs("NumInSale") & ")"
					.Write "{C}"
					If myApp.olkItemReport2 = "D" Then .Write rs("Currency") & " " & rs("Price") Else .Write rs("Price")
				End With
				WriteSep = True
			rs.movenext
			loop
		Else
			Response.Write "nodata"
		End If
	Case "CR" 'CommitedReport
		Source = Request("Source")
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKItemDetailsFetchOLKCommited" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@ItemCode") = ItemCode
		cmd("@Source") = Source
		If Not myAut.HasAuthorization(97) Then cmd("@SlpCode") = Session("vendid")
		cmd("@PriceDec") = myApp.PriceDec
		rs.open cmd, , 3, 1
		
		If Not rs.Eof Then
			WriteSep = False
			do while not rs.eof
				With Response
					If WriteSep Then Response.Write "{S}"
					.Write rs("ObjectCode")
					.Write "{C}"
					.Write rs("LogNum")
					.Write "{C}"
					.Write rs("LineNum")
					.Write "{C}"
					.Write rs("DocNum") 
					If Source = "O" Then Response.Write " (" & (CInt(rs("LineNum"))+1) & ")"
					.Write "{C}"
					.Write rs("ObjectCodeType")
					.Write "{C}"
					.Write FormatDate(rs("DocDate"), true)
					.Write "{C}"
					.Write rs("SlpName") 
					.Write "{C}"
					.Write rs("CardCode")
					.Write "{C}"
					.Write rs("CardName")
					.Write "{C}"
					 If rs("SaleType2") = 3 Then .Write FormatNumber(CDbl(rs("Quantity"))/CDbl(rs("SalPackUn")),0) Else .Write FormatNumber(CDbl(rs("Quantity")), myApp.QtyDec)
					.Write "{C}"
					.Write rs("SaleType")
					If (rs("SaleType2") = 2 or rs("SaleType2") = 3) and myApp.GetShowQtyInUn Then Response.Write "(" & rs("SaleTypeNum") & ")"
					If Not myApp.UnEmbPriceSet And rs("SaleType2") = 3 Then Response.Write " " & rs("SalUnitMsr")
					If myApp.GetShowQtyInUn Then Response.Write "(" & rs("NumInSale") & ")"
					.Write "{C}"
					
					If myApp.olkItemReport2 = "D" Then
						.Write rs("Currency") & " "
						If myApp.UnEmbPriceSet And rs("SaleType2") = 3 Then
							.Write FormatNumber(CDbl(rs("Price"))*CDbl(rs("SalPackUn")),myApp.PriceDec)
						Else 
							.Write FormatNumber(rs("Price"),myApp.PriceDec)
						End If 
					Else 
						.Write rs("Price")
					End If
				End With
				WriteSep = True
			rs.movenext
			loop
		Else
			Response.Write "nodata"
		End If
	Case "VD" 'Volume Discount
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKGetItemVolDiscData" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@ItemCode") = ItemCode
		cmd("@CardCode") = Session("UserName")
		cmd("@PriceList") = Session("PriceList")
		cmd("@Date") = SaveCmdDate(Request("Date"))
		cmd("@SaleType") = CInt(Request("SaleType"))
		cmd("@UnEmbPriceSet") = GetYN(myApp.UnEmbPriceSet)
		set rs = Server.CreateObject("ADODB.RecordSet")
		rs.open cmd, , 3, 1
		do while not rs.eof
			If rs.bookmark > 1 Then Response.Write "{S}"
			With Response
				.Write rs("Amount")
				.Write "{C}"
				.Write FormatNumber(rs("Price"), myApp.PriceDec)
			End With
		rs.movenext
		loop
	Case "CF" 'Change Field
		Quantity = CDbl(getNumericOut(Request("Quantity")))
		ManPrc = Request("ManPrc")
		SaleType = CInt(Request("SaleType"))
		FieldID = CInt(Request("FieldID"))
		
		If Quantity > 99999999.999999 Then Quantity = 99999999.999999
		
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKChangeItemDetailsFields" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LogNum") = Session("RetVal")
		cmd("@FieldID") = FieldID
		cmd("@ItemCode") = ItemCode
		cmd("@Quantity") = Quantity
		cmd("@SaleType") = SaleType
		cmd("@PrevSaleType") = CInt(Request("PrevSaleType"))
		cmd("@WhsCode") = Request("WhsCode")
		cmd("@Price") = CDbl(getNumericOut(Request("Price")))
		If Request("DiscPrcnt") <> "" Then cmd("@DiscPrcnt") = CDbl(getNumericOut(Request("DiscPrcnt")))
		cmd("@CardCode") = Session("UserName")
		cmd("@PriceList") = Session("PriceList")
		cmd("@ManPrc") = ManPrc
		cmd("@UnEmbPriceSet") = GetYN(myApp.UnEmbPriceSet)
		cmd("@Currency") = Request("Currency")
		set rs = cmd.execute()
		
		If SaleType = 3 Then FormatQty = 0 Else FormatQty = myApp.QtyDec
		
		Response.Write FormatNumber(CDbl(rs("Quantity")), FormatQty) 
		Response.Write "{S}" 
		Response.Write FormatNumber(CDbl(rs("Price")), myApp.PriceDec) & "{S}"
		If Not IsNull(rs("DiscPrcnt")) Then Response.Write FormatNumber(CDbl(rs("DiscPrcnt")), myApp.PercentDec)
		Response.Write "{S}"
		Response.Write FormatNumber(CDbl(rs("LineTotal")), myApp.SumDec)
		Response.Write "{S}"
		Response.Write rs("ChkInv")
		Response.Write "{S}"
		If Not IsNull(rs("Currency")) Then Response.Write rs("Currency") Else Response.Write myApp.MainCur
		Response.Write "{S}"		
		If (FieldID = 1 or FieldID = 2 or Request("OLKCombo") = "Y" and Request("Virtual") = "Y") and Request("isComp") <> "Y" Then
			CheckDiscount Request("TreeType"), Request("WhsCode"), Quantity, SaleType, Request("OLKCombo"), Request("ShowComp"), Request("Virtual"), CDbl(rs("DiscPrcnt"))
		End If
		Response.Write "{S}"
		If Request("Virtual") = "Y" Then Response.Write FormatNumber(VirtualTotal, myApp.SumDec)
		Response.Write "{S}"
		If Request("Virtual") = "Y" Then Response.Write VirtualChkInv
		Response.Write "{S}"
		If Request("Virtual") = "Y" Then Response.Write FormatNumber(CDbl(getNumericOut(Request("DiscPrcnt"))), myApp.PercentDec)
		Response.Write "{S}"
		If Request("Virtual") = "Y" Then Response.Write FormatNumber(VirtualTotal/Quantity, myApp.PriceDec)
	Case "CT"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKGetItemDetailsCompSum" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LogNum") = Session("RetVal")
		cmd("@UnEmbPriceSet") = GetYN(myApp.UnEmbPriceSet)
		cmd("@SumDec") = myApp.SumDec
		cmd("@PriceDec") = myApp.PriceDec
		set rs = Server.CreateObject("ADODB.RecordSet")
		
		priceTotal = 0.0
		compTotal = 0.0
		Quantity = CDbl(getNumericOut(Request("Quantity")))
		DiscPrcnt = CDbl(getNumericOut(Request("DiscPrcnt")))
		
		arrComp = Split(Request("CompData"), "{I}")
		For i = 0 to UBound(arrComp)
			arrData = Split(arrComp(i), "{S}")
			linePrice = CDbl(getNumericOut(arrData(0)))
			lineCur = arrData(1)
			lineQty = CDbl(getNumericOut(arrData(2)))
			lineUnit = CInt(arrData(3))
			lineItem = arrData(4)
			
			cmd("@ItemCode") = lineItem
			cmd("@Price") = linePrice
			cmd("@Cur") = lineCur
			cmd("@Qty") = lineQty
			cmd("@Unit") = lineUnit
			set rs = cmd.execute()

			priceTotal = priceTotal + CDbl(rs("PriceTotal"))			
			compTotal = compTotal + CDbl(rs("SumTotal"))
		Next
		
		If DiscPrcnt = 0 Then UnitPrice = priceTotal Else UnitPrice = priceTotal*100/DiscPrcnt
		With Response
			.Write FormatNumber(compTotal, myApp.SumDec)
			.Write "{S}"
			.Write FormatNumber(priceTotal/Quantity, myApp.PriceDec)
			.Write "{S}"
			.Write FormatNumber(UnitPrice, myApp.PriceDec)
		End With
	Case "OCO", "OCS"
		mySource = "O"
		If DataType = "OCS" Then mySource = "S"
		
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKItemDetailsFetchOLKCommited" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@ItemCode") = ItemCode
		cmd("@Source") = mySource
		If Not myAut.HasAuthorization(97) Then cmd("@SlpCode") = Session("vendid")
		cmd("@PriceDec") = myApp.PriceDec
		set rs = cmd.execute()
		If Not rs.Eof Then
			do while not rs.eof
				For i = 0 to rs.Fields.Count -1
					Response.Write rs(i) & "{C}"
				Next
				Response.Write "{S}"
			rs.movenext
			loop
		Else
			Response.Write "nodata"
		End If
	Case "BP"
		If Session("UserName") <> "" Then
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKItemDetailsFetchBestPrice" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@CardCode") = Session("UserName")
			cmd("@ItemCode") = ItemCode
			If not myAut.HasAuthorization(97) Then cmd("@SlpCode") = Session("vendid")
			set rs = cmd.execute()
			If Not rs.Eof Then
				Response.Write "ok|" & rs(0) & "&nbsp;" & FormatNumber(rs(1), myApp.PriceDec) & "|" & FormatNumber(rs(2), myApp.QtyDec) & "|"
				set cmd = Server.CreateObject("ADODB.Command")
				cmd.ActiveConnection = connCommon
				cmd.CommandType = &H0004
				cmd.CommandText = "DBOLKItemDetailsFetchBestPriceHistory" & Session("ID")
				cmd.Parameters.Refresh()
				cmd("@CardCode") = Session("UserName")
				cmd("@ItemCode") = ItemCode
				If not myAut.HasAuthorization(97) Then cmd("@SlpCode") = Session("vendid")
				rs.close
				rs.open cmd, , 3, 1
				do while not rs.eof
					If rs.bookmark > 1 Then Response.Write "{S}"
					Response.Write rs("DocEntry") & "{C}" & rs("DocNum") & "{C}" & rs("LineNum") & "{C}" & rs("Currency") & "&nbsp;" & FormatNumber(rs("Price"), myApp.PriceDec) & "{C}" & FormatDate(rs("DocDate"), True) & "{C}" & FormatNumber(rs("Quantity"), myApp.PriceDec)
				rs.movenext
				loop
			Else
				Response.Write "nodata"
			End If
		Else
			Response.Write "nodata"
		End If
	Case "IR"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKItemDetailsFetchCustomAgentData" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@ItemCode") = ItemCode
		cmd("@branchIndex") = Session("branch")
		If Session("PriceList") <> "" Then cmd("@PriceList") = Session("PriceList")
		cmd("@CardCode") = Session("UserName")
		cmd("@SlpCode") = Session("vendid")
		If Request("WhsCode") <> "" Then cmd("@WhsCode") = Request("WhsCode")
		If Request("ItemCmd") = "A" Then
			cmd("@Quantity") = CDbl(getNumericOut(Request("Quantity")))
			cmd("@Unit") = Request("SaleType")
			cmd("@Price") = CDbl(getNumericOut(Request("Price")))
		End If
		set rs = cmd.execute()
		For each fld in rs.Fields
			Response.Write Replace(fld.Name, "-", "_") & "{C}" & fld
			Response.Write "{S}"
		Next
	Case "IRD"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKGetItemRepLinkVals" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@rowIndex") = CInt(Request("rowIndex"))
		cmd("@rsIndex") = CInt(Request("rsIndex"))
		If Session("PriceList") <> "" Then cmd("@PriceList") = Session("PriceList")
		cmd("@SlpCode") = Session("vendid")
		If Session("UserName") <> "" Then cmd("@CardCode") = Session("UserName")
		If Request("WhsCode") <> "" Then cmd("@WhsCode") = Request("WhsCode")
		cmd("@ItemCode") = ItemCode
		set rs = cmd.execute()
		do while not rs.eof
			With Response
				.Write rs(0)
				.Write "{C}"
				.Write rs(1)
				.Write "{C}"
				.Write rs(2)
				.Write "{C}"
				If Not IsNull(rs(3)) Then .Write rs(3)
				.Write "{C}"
				If Not IsNull(rs(4)) Then .Write FormatDate(rs(4), False)
				.Write "{C}"
				.Write rs(5)
				.Write "{S}"
			End With
		rs.movenext
		loop
	Case "Inv"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKSalesItemDetails" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@ItemCode") = ItemCode
		cmd("@WhsCode") = Request("WhsCode")
		cmd("@LanID") = Session("LanID")
		set rs = cmd.execute()
		
		Response.Write GetYN(RS("ORDRChk") ="True" and myAut.HasAuthorization(103))
		Response.Write "|"
		Response.Write GetYN(RS("ObsChk") = "True" and myAut.HasAuthorization(103))
		Response.Write "|"
		Response.Write FormatNumber(rs("OnHand"), 0)
		Response.Write "|"
		Response.Write FormatNumber(RS("OnHandUnVentSAP"),0) & " (" & FormatNumber(RS("OnHandSueltoUnVentaSAP"),0) & ")"
		Response.Write "|"
		Response.Write FormatNumber(RS("OnHandUnEmbSAP"),0) & " (" & FormatNumber(RS("OnHandSueltoUnEmbSAP"),0) & ")"
		Response.Write "|"
		Response.Write FormatNumber(RS("DispSAP"),0)
		Response.Write "|"
		Response.Write FormatNumber(RS("DispUnVentSAP"),0) & " (" & FormatNumber(RS("DispSueltoUnVentSAP"),0) & ")"
		Response.Write "|"
		Response.Write FormatNumber(RS("DispUnEmbSAP"),0) & " (" & FormatNumber(RS("DispSueltoUnEmbSAP"),0) & ")"
		Response.Write "|"
		Response.Write FormatNumber(RS("InvBDGWhs"),0)
		Response.Write "|"
		Response.Write FormatNumber(RS("InvUnVentBDGWhs"),0) & " (" & FormatNumber(RS("InvSueltoUnVentBDGWhs"),0) & ")"
		Response.Write "|"
		Response.Write FormatNumber(RS("InvUnEmbBDGWhs"),0) & " (" & FormatNumber(RS("InvSueltoUnEmbBDGWhs"),0) & ")"
		Response.Write "|"
		Response.Write FormatNumber(RS("InvBDGDisp"),0)
		Response.Write "|"
		Response.Write FormatNumber(RS("InvBDGUnVentDisp"),0) & " (" & FormatNumber(RS("InvBDGSueltoUnVentDisp"),0) & ")"
		Response.Write "|"
		Response.Write FormatNumber(RS("InvBDGUnEmbDisp"),0) & " (" & FormatNumber(RS("InvBDGSueltoUnEmbDisp"),0) & ")"
		Response.Write "|"
		Response.Write FormatNumber(RS("InvOLKDisp"),0)
		Response.Write "|"
		Response.Write FormatNumber(RS("InvOLKUnVentDisp"),0) & " (" & FormatNumber(RS("InvOLKSueltoUnVentDisp"),0) & ")"
		Response.Write "|"
		Response.Write FormatNumber(RS("InvOLKUnEmbDisp"),0) & " (" & FormatNumber(RS("InvOLKSueltoUnEmbDisp"),0) & ")"
		Response.Write "|"
		Response.Write FormatNumber(RS("InvOLKBDGDisp"),0)
		Response.Write "|"
		Response.Write FormatNumber(RS("InvOLKBDGUnVentDisp"),0) & " (" & FormatNumber(RS("InvOLKBDGSueltoUnVentDisp"),0) & ")"
		Response.Write "|"
		Response.Write FormatNumber(RS("InvOLKBDGUnEmbDisp"),0) & " (" & FormatNumber(RS("InvOLKBDGSueltoUnEmbDisp"),0) & ")"
	Case "L" 'Load 
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKGetWarehouses" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		rs.open cmd, , 3, 1
		do while not rs.eof
			If rs.bookmark > 1 Then Response.Write "{O}"
			Response.Write rs(0) & "{C}" & rs(1)
		rs.movenext
		loop
		
		Response.Write "{S}"
		
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKGetItemRepRead" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@Access") = "V"
		cmd("@OP") = "O"
		cmd("@UserAccess") = Session("UserAccess")
		cmd("@SlpCode") = Session("vendid")
		If Session("PriceList") = "" Then cmd("@FilterPriceList") = "Y"
		If Request("ItemCmd") <> "A" Then cmd("@FilterCartVars") = "Y"
		If Session("UserAccess") = "U" Then cmd("@rgIndex") = myAut.AuthorizedRepGroups
		rs.close
		rs.open cmd, , 3, 1
		do while not rs.eof
			If rs.bookmark > 1 Then Response.Write "{O}"
			Response.Write rs(0)
			Response.Write "{C}"
			Response.Write rs(1)
			Response.Write "{C}"
			Response.Write rs(2)
			Response.Write "{C}"
			Response.Write rs(3)
			Response.Write "{C}"
			Response.Write rs(4)
			Response.Write "{C}"
			Response.Write rs(5)
		rs.movenext
		loop
		
		Response.Write "{S}"
		
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKGetItemDispORDR" & Session("ID")
		cmd.Parameters.Refresh()
		rs.close
		rs.open cmd, , 3, 1
		do while not rs.eof
			If rs.bookmark > 1 Then Response.Write "{C}"
			Response.Write rs(0)
		rs.movenext
		loop
		
		Response.Write "{S}"
		Response.Write myApp.VerfyDisp & myApp.VerfyDispWhs
		
		Response.Write "{S}"
		If myAut.HasAuthorization(68) Then Response.Write GetYN(myApp.ShowLineDiscount) Else Response.Write "N"
		
		Response.Write "{S}"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKGetLineNotesList" & Session("ID")
		cmd.Parameters.Refresh()
		rs.close
		rs.open cmd, , 3, 1
		do while not rs.eof
			If rs.bookmark > 1 Then Response.Write "{N}"
			Response.Write rs(0) & "{C}" & rs(1)
		rs.movenext
		loop
		
		Response.Write "{S}"
		Response.Write GetYN(myApp.SDKLineMemo)
		Response.Write "{S}"
		If Not myAut.HasAuthorization(99) Then Response.Write "Y" Else Response.Write "N"
		
		Response.Write "{S}"
		If myApp.LawsSet = "MX" or myApp.LawsSet = "CL" or myApp.LawsSet = "CR" or myApp.LawsSet = "GT" or myApp.LawsSet = "US" or myApp.LawsSet = "CA" Then
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKGetAddItemTaxCodeList" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@LanID") = Session("LanID")
			set rd = Server.CreateObject("ADODB.RecordSet")
			rd.open cmd, , 3, 1
			do while not rd.eof 
				If rd.bookmark > 1 Then Response.Write "{C}"
				Response.Write rd("Code") & "{D}" & rd("Name")
			rd.movenext
			loop
		End If
	Case "AI"
		AddItem	
	Case "AWL"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKItemAddToWL" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@CardCode") = Session("UserName")
		cmd("@ItemCode") = ItemCode
		cmd.execute()
		Response.Write "ok"
End Select

Sub AddItem
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandText = "DBOLKItemAddToCart" & Session("ID")
	cmd.CommandType = &H0004
	cmd.Parameters.Refresh
	cmd("@LogNum") = Session("RetVal")
	cmd("@branch") = Session("branch")
	cmd("@UserType") = userType
	cmd("@SlpCode") = Session("vendid")
	cmd("@CardCode") = Session("UserName")
	cmd("@PriceList") = Session("PriceList")
	cmd("@LawsSet") = myApp.LawsSet
	cmd("@BasketMItems") = GetYN(myApp.BasketMItems)
	cmd("@EnableCItemPurLog") = GetYN(myApp.EnableCItemPurLog)
	cmd("@All") = Request("AddAll")
			
	Qty = CDbl(getNumericOut(Request("Qty")))
	SaleType = CInt(Request("SaleType"))
	DiscPrcnt = CDbl(getNumericOut(Request("DiscPrcnt")))
	Price = CDbl(getNumericOut(Request("Price")))
	ManPrc = Request("ManPrc")
	TreeType = Request("TreeType")
	Whs = Request("Whs")
	Note = Request("Note")
	CompType = Request("CompType")
	Virtual = Request("Virtual")
	HideComp = Request("HideComp")
	CompData = Request("CompData")
	TaxCode = Request("TaxCode")
	itmCurrency = Request("Currency")
	
	IgnorePrice = CompType = "SaleTree" and HideComp = "N" 
	If IgnorePrice Then 
		Price = 0
		ManPrc = "Y"
	End If
	
	cmd("@ItemCode") = ItemCode
	If SaleType <> "" Then cmd("@SaleType") = SaleType
	If Qty <> "" Then cmd("@Qty") = Qty
	If Price <> "" Then cmd("@Price") = Price
	If CompType = "OLKCombo" and Virtual = "Y" Then 
		cmd("@VirtualDiscPrcnt") = DiscPrcnt
	End If
	If ManPrc <> "" Then cmd("@ManPrc") = ManPrc
	If CompType = "SaleTree" Then cmd("@TreeType") = "S"
	If CompType = "OLKCombo" and Virtual = "Y" Then cmd("@VirtualTreeType") = "S"
	
	If Whs <> "" Then cmd("@WhsCode") = Whs
	
	If TaxCode <> "" Then cmd("@TaxCode") = TaxCode
	If Note <> "" Then cmd("@Note") = Note
	cmd("@Currency") = itmCurrency
	
	cmd.execute()
	
	strRetVal = cmd("@RetVal")
	FatherLineID = cmd("@LineNum")
	
	Response.Write strRetVal & "{S}"
	
	If CompType <> "" Then
		ArrComp = Split(CompData, "{I}")
		For i = 0 to UBound(ArrComp)
			ArrCompData = Split(ArrComp(i), "{S}")
			compItem = ArrCompData(0)
			compQty = CDbl(getNumericOut(ArrCompData(1)))
			compUnit = ArrCompData(2)
			compWhs = ArrCompData(3)
			compPrice = CDbl(getNumericOut(ArrCompData(4)))
			compDiscPrcnt = CDbl(getNumericOut(ArrCompData(10)))
			compManPrc = ArrCompData(5)
			compTaxCode = ArrCompData(6)
			compRecQty = CDbl(getNumericOut(ArrCompData(7)))
			compHideComp = ArrCompData(8)
			compChildID = ArrCompData(9)
			
			IgnorePrice = CompType = "SaleTree" and compHideComp = "Y" 
			virtualPrice = 0
			If CompType = "OLKCombo" and Virtual = "Y" Then
				virtualPrice = compPrice
			End If
			
			If IgnorePrice Then 
				compPrice = 0
				compManPrc = "Y"
			End If
		
			cmd("@ItemCode") = compItem
			If compUnit <> "" Then cmd("@SaleType") = compUnit
			If compQty <> "" Then cmd("@Qty") = compQty
			If CompType = "OLKCombo" and Virtual = "Y" Then 
				cmd("@VirtualDiscPrcnt") = compDiscPrcnt
			End If
			If compPrice <> "" Then cmd("@Price") = compPrice
			If compManPrc <> "" Then cmd("@ManPrc") = compManPrc
			Select Case CompType
				Case "SaleTree", "OLKCombo"
					If CompType = "SaleTree" Then cmd("@TreeType") = "C"
					cmd("@FatherLineID") = FatherLineID
					cmd("@Father") = ItemCode
			End Select
			If compWhs <> "" Then cmd("@WhsCode") = compWhs
			
			If compTaxCode <> "" Then cmd("@TaxCode") = compTaxCode
			
			cmd("@RecQty") = compRecQty
			
			If compChildID <> "" Then
				cmd("@VirtualChildID") = CInt(compChildID)
				If CompType = "OLKCombo" and Virtual = "Y" Then
					cmd("@VirtualTreeType") = "C"
					cmd("@VirtualPrice") = virtualPrice
				End If
			End If
			
			cmd.execute()
			
			strRetVal = cmd("@RetVal")
			
			Response.Write strRetVal & "{S}"
		Next
		
		If CompType = "OLKCombo" and Virtual = "Y" Then
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKItemAddToCartEnd" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@LogNum") = Session("RetVal")
			cmd("@LineNum") = FatherLineID
			cmd.execute()
		End If
	End If
	
End Sub

Sub CheckDiscount(ByVal TreeType, ByVal WhsCode, ByVal Quantity, ByVal SaleUnit, ByVal OLKCombo, ByVal ShowComp, ByVal Virtual, ByVal Discount)
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	itemType = ""
	If TreeType = "S" Then
		Response.Write "SaleTree{C}"
		cmd.CommandText = "DBOLKGetItemTreeComp" & Session("ID")
		itemType = "T"
	ElseIf OLKCombo = "Y" Then
		Response.Write "OLKCombo{C}"
		cmd.CommandText = "DBOLKGetItemComboComp" & Session("ID")
		itemType = "C"
	ElseIf myApp.EnableItemRec Then
		cmd.CommandText = "DBOLKGetItemRecByQuery"
		itemType = "R"
	Else
		Exit Sub
	End If

	cmd.Parameters.Refresh()
	If itemType = "R" Then cmd("@dbID") = Session("ID")
	cmd("@LanID") = Session("LanID")
	cmd("@branch") = Session("branch")
	If Request("cmd") = "A" Then 
		cmd("@LogNum") = Session("RetVal")
		cmd("@CardCode") = Session("UserName")
		cmd("@PriceList") = Session("PriceList")
	End If
	If itemType = "C" and OLKCombo = "Y" and Virtual = "Y" and Discount <> 0 Then cmd("@Discount") = Discount
	
	cmd("@ItemCode") = ItemCode
	cmd("@SlpCode") = Session("vendid")
	cmd("@WhsCode") = WhsCode
	cmd("@TreePricOn") = GetYN(myApp.TreePricOn)
	cmd("@UnEmbPriceSet") = GetYN(myApp.UnEmbPriceSet)
	cmd("@SaleType") = SaleUnit
	cmd("@Quantity") = Quantity
	cmd("@QtyDec") = myApp.QtyDec
	set rd = Server.CreateObject("ADODB.RecordSet")
	rd.open cmd, , 3, 1
	strRec = False
	If Not rd.Eof Then
		If TreeType <> "S" and OLKCombo <> "Y" and myApp.EnableItemRec Then Response.Write "ItemRec{C}"
		do while not rd.eof
			If strRec Then Response.Write "{R}"
			Response.Write rd("ItemCode")
			Response.Write "{O}"
			Response.Write rd("ItemName")
			Response.Write "{O}"
			Response.Write FormatNumber(CDbl(rd("Quantity")), FormatQty)
			Response.Write "{O}"
			Response.Write FormatNumber(CDbl(rd("Price")), myApp.PriceDec)
			Response.Write "{O}"
			Response.Write rd("PicturName")
			Response.Write "{O}"
			Response.Write rd("DocEntry")
			Response.Write "{O}"
			Response.Write rd("Currency")
			Response.Write "{O}"
			Response.Write rd("Checked")
			Response.Write "{O}"
			Response.Write rd("Locked")
			Response.Write "{O}"
			Response.Write rd("WhsCode")
			Response.Write "{O}"
			Response.Write rd("Comment")
			Response.Write "{O}"
			Response.Write rd("HideComp")
			Response.Write "{O}"
			Response.Write rd("SaleTypeDesc")
			Response.Write "{O}"
			Response.Write FormatNumber(CDbl(rd("LineTotal")), myApp.SumDec)
			Response.Write "{O}"
			Response.Write rd("DocCur")
			Response.Write "{O}"
			Response.Write rd("DiscType")
			Response.Write "{O}"
			Response.Write FormatNumber(CDbl(rd("Discount")), myApp.PercentDec)
			Response.Write "{O}"
			Select Case itemType
				Case "T"
					Response.Write "Y"
				Case "C"
					Response.Write rd("LockDisc")
				Case "R"
					Response.Write "N"
			End Select
			Response.Write "{O}"
			Response.Write rd("NumInSale")
			Response.Write "{O}"
			Response.Write rd("SalUnitMsr")
			Response.Write "{O}"
			Response.Write rd("SalPackUn")
			Response.Write "{O}"
			Response.Write rd("SalPackMsr")
			Response.Write "{O}"
			Response.Write rd("chkInv")
			Response.Write "{O}"
			Response.Write rd("TreeType")
			Response.Write "{O}"
			Response.Write rd("ChkVolDisc")
			Response.Write "{O}"
			Response.Write FormatNumber(CDbl(rd("RecQty")), myApp.QtyDec)
			Response.Write "{O}"
			If OLKCombo = "Y" Then Response.Write rd("LockQty") Else Response.Write rd("Locked")
			strRec = True
			If Virtual = "Y" Then
				VirtualTotal = VirtualTotal + CDbl(rd("LineTotal"))
				If rd("chkInv") = "N" Then VirtualChkInv = "N"
				Response.Write "{O}"
				Response.Write rd("CompID")
			End If
		rd.movenext
		loop
	End If
End Sub
%>
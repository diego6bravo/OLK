<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

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

LogNum = Session("RetVal")
LineNum = CInt(Request("LineNum"))
ProcType = Request.Form("ProcType")
SumDec = myApp.SumDec

set rs = Server.CreateObject("ADODB.RecordSet")
Select Case ProcType
	Case "Quantity"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKCartSaveLineQty" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LogNum") = LogNum
		cmd("@LineNum") = LineNum
		cmd("@Quantity") = CDbl(getNumericOut(Request.Form("Quantity")))
		cmd("@PriceList") = Session("PriceList")
		cmd("@UnEmbPriceSet") = GetYN(myApp.UnEmbPriceSet)
		cmd("@PriceDec") = myApp.PriceDec
		cmd("@PercentDec") = myApp.PercentDec
		cmd("@SumDec") = myApp.SumDec
		cmd("@UserType") = userType
		cmd.execute()
		If cmd("@FormatQty") = "Y" Then
			Response.Write FormatNumber(CDbl(cmd("@Quantity")), myApp.QtyDec) 
		Else
			Response.Write FormatNumber(CDbl(cmd("@Quantity")), 0) 
		End If
		Response.Write "{S}" 
		Response.Write FormatNumber(CDbl(cmd("@Price")), myApp.PriceDec)
		Response.Write "{S}"
		Response.Write cmd("@Currency") 
		Response.Write "{S}" 
		Response.Write FormatNumber(CDbl(cmd("@DiscPrcnt")), myApp.PercentDec)
		Response.Write "{S}" 
		Response.Write FormatNumber(CDbl(cmd("@LineTotal")), SumDec)
		Response.Write "{S}" 
		Response.Write cmd("@ErrOfferQty") 
		Response.Write "{S}" 
		Response.Write cmd("@OfertIndex") 
		Response.Write "{S}" 
		Response.Write cmd("@UpdChilds")
		Response.Write "{S}" 
		Response.Write cmd("@ChkInv")
		Response.Write "{S}" 
		Response.Write cmd("@LockAdd")
		
		If cmd("@UpdChilds") = "Y" Then
			Response.Write "{S}"
			
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKGetCartLineChildData" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@LogNum") = LogNum
			cmd("@LineNum") = LineNum
			cmd("@SumDec") = myApp.SumDec
			
			set rs = cmd.execute()
			do while not rs.eof
				Response.Write "{L}" & rs(0) & "{C}" & FormatNumber(CDbl(rs(1)), myApp.QtyDec) & "{C}" & FormatNumber(CDbl(rs(2)), SumDec)
			rs.movenext
			loop
		End If
	Case "SaleType"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKCartSaveLineSaleType" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LogNum") = LogNum
		cmd("@LineNum") = LineNum
		cmd("@Quantity") = CDbl(getNumericOut(Request.Form("Quantity")))
		cmd("@SaleType") = CInt(Request("SaleType"))
		cmd("@PriceList") = Session("PriceList")
		cmd("@UnEmbPriceSet") = GetYN(myApp.UnEmbPriceSet)
		cmd("@PriceDec") = myApp.PriceDec
		cmd("@PercentDec") = myApp.PercentDec
		cmd("@SumDec") = myApp.SumDec
		cmd("@UserType") = userType
		cmd.execute()
		If cmd("@FormatQty") = "Y" Then
			Response.Write FormatNumber(CDbl(cmd("@Quantity")), myApp.QtyDec) 
		Else
			Response.Write FormatNumber(CDbl(cmd("@Quantity")), 0) 
		End If
		Response.Write "{S}" 
		Response.Write FormatNumber(CDbl(cmd("@Price")), myApp.PriceDec)
		Response.Write "{S}" 
		Response.Write FormatNumber(CDbl(cmd("@UnitPrice")), myApp.PriceDec)
		Response.Write "{S}"
		Response.Write cmd("@Currency") 
		Response.Write "{S}" 
		Response.Write FormatNumber(CDbl(cmd("@DiscPrcnt")), myApp.PercentDec)
		Response.Write "{S}" 
		Response.Write FormatNumber(CDbl(cmd("@LineTotal")), SumDec)
		Response.Write "{S}"  
		Response.Write cmd("@ChkInv")
		Response.Write "{S}" 
		Response.Write cmd("@LockAdd")
	Case "Discount"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKCartSaveLineDiscount" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LogNum") = LogNum
		cmd("@LineNum") = LineNum
		cmd("@DiscPrcnt") = CDbl(getNumericOut(Replace(Request.Form("Discount"), "%", "")))
		cmd("@PriceList") = Session("PriceList")
		cmd("@UnEmbPriceSet") = GetYN(myApp.UnEmbPriceSet)
		cmd("@PriceDec") = myApp.PriceDec
		cmd("@PercentDec") = myApp.PercentDec
		cmd("@SumDec") = myApp.SumDec
		cmd("@SlpCode") = Session("vendid")
		cmd("@UserAccess") = Session("UserAccess")
		cmd.execute()
		Response.Write FormatNumber(CDbl(cmd("@Price")), myApp.PriceDec)
		Response.Write "{S}" 
		Response.Write cmd("@Currency") 
		Response.Write "{S}" 
		Response.Write FormatNumber(CDbl(cmd("@LineTotal")), SumDec)
		Response.Write "{S}"  
		Response.Write FormatNumber(CDbl(cmd("@DiscPrcnt")), myApp.PercentDec)
		Response.Write "{S}" 
		Response.Write cmd("@ErrMaxDisc")
	Case "Price"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKCartSaveLinePrice" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LogNum") = LogNum
		cmd("@LineNum") = LineNum
		cmd("@Price") = CDbl(getNumericOut(Replace(Request.Form("Price"), Request.Form("Cur"), "")))
		cmd("@PriceList") = Session("PriceList")
		cmd("@UnEmbPriceSet") = GetYN(myApp.UnEmbPriceSet)
		cmd("@PriceDec") = myApp.PriceDec
		cmd("@PercentDec") = myApp.PercentDec
		cmd("@SumDec") = myApp.SumDec
		cmd("@SlpCode") = Session("vendid")
		cmd("@UserAccess") = Session("UserAccess")
		cmd.execute()
		Response.Write FormatNumber(CDbl(cmd("@Price")), myApp.PriceDec)
		Response.Write "{S}" 
		Response.Write cmd("@Currency") 
		Response.Write "{S}" 
		Response.Write FormatNumber(CDbl(cmd("@LineTotal")), SumDec)
		Response.Write "{S}"  
		Response.Write FormatNumber(CDbl(cmd("@DiscPrcnt")), myApp.PercentDec)
		Response.Write "{S}" 
		Response.Write cmd("@ErrMaxDisc")
End Select

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetDocTotalData" & Session("ID")
cmd.Parameters.Refresh()
cmd("@LogNum") = LogNum
If Session("PayCart") Then cmd("@MC") = "Y"
set rs = Server.CreateObject("ADODB.RecordSet")
rs.open cmd, , 3, 1

Response.Write "{T}" & FormatNumber(CDbl(rs(0)), SumDec) & "{S}" & FormatNumber(CDbl(rs(1)), SumDec) & "{S}" & FormatNumber(CDbl(rs(2)), SumDec) & "{S}" & FormatNumber(CDbl(rs(3)), SumDec) & "{S}" & FormatNumber(CDbl(rs(4)), SumDec) & "{S}" & FormatNumber(CDbl(rs(5)), SumDec) 
If Session("PayCart") Then Response.Write "{S}" & FormatNumber(CDbl(rs(5)), SumDec)
			
 %>
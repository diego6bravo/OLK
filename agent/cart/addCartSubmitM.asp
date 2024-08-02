<%@ Language=VBScript%>
<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="../chkLogin.asp" -->
<!--#include file="../myHTMLEncode.asp" -->
<!--#include file="../lcidReturn.inc" -->
<%
Dim Curr
If userType = "C" Then MainDoc = "default.asp" Else MainDoc = "agent.asp"
errMsg = ""
fastAddErr = ""
If Request("SaleType") <> "" Then SaleType = Request("SaleType") Else SaleType = "NULL"
If Request("precio") <> "" Then Precio = getNumeric(Request("Precio")) Else Precio = "NULL"
If Request("WhsCode") <> "" Then WhsCode = "N'" & saveHTMLDecode(Request("WhsCode"), False) & "'" Else WhsCode = "OLKCommon.dbo.DBOLKGetWhsCode" & Session("ID") & "(" & Session("branch") & ", " & Session("vendid") & ", @ItemCode)"

EnableItemRec = myApp.EnableItemRec
LogPur = myApp.EnableCItemPurLog and userType = "C"
ItemCode = Request("Item")

 	
If Request("fastAdd") = "Y" Then
	If myApp.FastAddUnRem Then
		Session("CurSaleType") = SaleType
	End If

	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKCheckFastAddItem" & Session("ID")
	cmd.Parameters.Refresh() 
	cmd("@UserType") = userType
	cmd("@UserAccess") = Session("UserAccess")
	cmd("@SlpCode") = Session("vendid")
	cmd("@CardCode") = Session("UserName")
	cmd("@ItemCode") = ItemCode
	If Request("precio") <> "" Then cmd("@Price") = CDbl(getNumericOut(Request("precio")))
	If Request("SaleType") <> "" Then cmd("@SaleType") = Request("SaleType")
	cmd("@PriceList") = Session("PriceList")
	If myApp.GetApplyGenFilter Then cmd("@ApplyGenFilter") = "Y"
	cmd.execute()
	
	Select Case cmd("@Check")
		Case 0
			ItemCode = cmd("@ItemCode")
		Case 1
			fastAddErr = "Y"
			errMsg = errMsg & "&fastAddErr=Y&fastAddErrItm=" & saveHTMLDecode(ItemCode, False)
		Case 2
			fastAddErr = "Y"
			ItemCode = cmd("@ItemCode")
			errMsg = errMsg & "&fastAddErr=Y&fastAddErrType=D&fastAddErrItm=" & saveHTMLDecode(Request("Item"), False)
		Case 3, 4, 5, 6
			ItemCode = cmd("@ItemCode")
			errMsg = errMsg & "&fastAddErr=Y&fastAddErrType=F&fastAddErrItm=" & saveHTMLDecode(Request("Item"), False)
			fastAddErr = "Y"
	End Select
End If
	      	
set rs = server.createobject("ADODB.RecordSet")
    
If AddErr = "" and fastAddErr = "" Then

	'Revisa si hay suficiente Stock para agregarlo al shopping cart
	set rv = Server.CreateObject("ADODB.RecordSet")
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKCheckInvBefAdd" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@LogNum") = Session("RetVal")
	cmd("@ItemCode") = ItemCode
	If Request("WhsCode") <> "" Then cmd("@WhsCode") = Request("WhsCode")
	If Request("SaleType") <> "" Then cmd("@SaleType") = Request("SaleType")
	cmd("@FirstQuantity") = CDbl(getNumericOut(Request("T1")))
	cmd("@branch") = Session("branch")
	cmd("@SlpCode") = Session("vendid")
	cmd("@UserType") = userType
	set rv = cmd.execute()
	Verfy = rv("Verfy") = "Y"
	ItmEntry = rv("DocEntry")
	TreeType = rv("TreeType")
	If TreeType = "S" Then EnableItemRec = False
	VerfyInCart = rv("VerfyInCart") = "True"
	lineNum = rv("VerfyInCartLineNum")
	If TreeType = "S" Then 
		EnableItemRec = False
		RecType = 2
	End If
	
	
	ItemHasRec = False
	
	If EnableItemRec Then
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKGetRecItems" & Session("ID")
		cmd("@ItemCode") = ItemCode
		cmd("@LogNum") = Session("RetVal")
		cmd("@CardCode") = Session("username")
		cmd("@SlpCode") = Session("vendid")
		If Request("WhsCode") <> "" Then cmd("@WhsCode") = Request("WhsCode")
		cmd("@branch") = Session("branch")
		cmd("@LanID") = Session("LanID")
		set rs = Server.CreateObject("ADODB.RecordSet")
		conn.execute("set ROWCOUNT 1")
		set rs = cmd.execute()
		ItemHasRec = Not rs.Eof
		RecType = 1
	End If	
	
	doGoRec = ItemHasRec or TreeType = "S"

		
	If Verfy and not doGoRec Then
		
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKCartAddSFM" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@lognum") = Session("RetVal")
		cmd("@SlpCode") = Session("vendid")
		
		AddItem = True
		If Request("SaleType") <> "" Then SaleType = Request("SaleType") Else SaleType = "NULL"
		If Request("OfertIndex") <> "" Then OfertIndex = Request("OfertIndex") Else OfertIndex = "NULL"
		If Session("branch") <> "" Then branch = Session("branch") Else branch = "-1"
		TaxCode = ""
	    If myApp.LawsSet = "MX" or myApp.LawsSet = "CL" or myApp.LawsSet = "CR" or myApp.LawsSet = "GT" or myApp.LawsSet = "US" or myApp.LawsSet = "CA" Then
	      	If Request("TaxCode") <> "" Then
		      	TaxCode = Request("TaxCode")
		    Else
		      	TaxCode = getTaxCode
		    End If
		    
	      	If TaxCode = "Disabled" Then 
	      		TaxCode = "NULL"
	      	ElseIf TaxCode = "" Then
      			errMsg = "&err=tax&tItem=" & Request("Item")
      			AddItem = False
      		End If
		End If
		If AddItem and TreeType <> "T" Then
			cmd("@Quantity") = CDbl(Request("T1"))
			cmd("@item") = ItemCode
			If Request("precio") <> "" Then cmd("@itemprice") = CDbl(Request("Precio"))
			cmd("@PriceList") = Session("PriceList")
			cmd("@UserType") = userType
			If SaleType  <> "NULL" Then cmd("@SaleType") = SaleType 
			cmd("@branchIndex") = Session("branch")
			If myApp.LawsSet = "MX" or myApp.LawsSet = "CL" or myApp.LawsSet = "CR" or myApp.LawsSet = "GT" or myApp.LawsSet = "US" or myApp.LawsSet = "CA" Then
				If TaxCode <> "NULL" Then cmd("@TaxCode") = TaxCode
			End If
			If Request("Locked") = "Y" Then cmd("@Locked") = "Y"
			cmd.execute()
				
			If Request("redir") = "wish" Then
				set cmd = Server.CreateObject("ADODB.Command")
				cmd.ActiveConnection = connCommon
				cmd.CommandType = &H0004
				cmd.CommandText = "DBOLKWLDel" & Session("ID")
				cmd.Parameters.Refresh()
				cmd("@CardCode") = Session("UserName")
				cmd("@ItemCode") = Request("Item")
				cmd.execute()
			End If
	      
			If LogPur Then
				set cmd = Server.CreateObject("ADODB.Command")
				cmd.ActiveConnection = connCommon
				cmd.CommandType = &H0004
				cmd.CommandText = "DBOLKAddItemLog" & Session("ID")
				cmd.Parameters.Refresh()
				cmd("@CardCode") = Session("UserName")
				cmd("@ItemCode") = Request("Item")
				cmd("@LineType") = "P"
				cmd.execute()
			End If
		End If
	ElseIf Not doGoRec Then
		errMsg = "&err=inv"
	End If
End If
Conn.Close
ClosePop = False

If Request("redir") = "report" Then
	redirVal = "../report.asp?" & Replace(Replace(Request("retVal"), "{y}", "&"), "{i}", "=") & errMsg
ElseIf myApp.GetAfterCartAdd = "Y" and Request("redir") <> "history" Then
	redirVal = "../cart.asp?cmd=" & Session("Cart") & errMsg
ElseIf Request("redir") = "no" Then 
	redirVal = "../search.asp?cmd=searchCart" & Replace(Replace(Request("retVal"), "{y}", "&"), "{i}", "=") & errMsg
ElseIf Request("redir") = "searchCashCart" Then 
	redirVal = "../search.asp?cmd=searchCashCart&page=" & Request("page") & "&document=" & Request("document") & errMsg 
ElseIf Request("redir") = "wish" Then
	redirVal = "../wish.asp?cmd=wish" & Replace(Replace(Request("retVal"), "{y}", "&"), "{i}", "=") & errMsg 
ElseIf Request("redir") = "home" then
	redirVal = "../" & MainDoc & "?cmd=home" & errMsg 
ElseIf Request("redir") = "cashInv" then
	redirVal = "../cart.asp?cmd=cashInv" & errMsg 
ElseIf Request("redir") = "prom" then
	redirVal = "../prom.asp?cmd=prom" & errMsg 
ElseIf Request("redir") = "clientProm" then
	redirVal = "../" & MainDoc & "?cmd=clientProm" & errMsg 
ElseIf Request("redir") = "cart" then
	If Request("Bookmark") <> "" Then Bookmark = "&#" & Request("BookMark")
	redirVal = "../cart.asp?cmd=cart&focus=" & Request("focus") & errMsg & BookMark 
ElseIf Request("redir") = "oferts" then
	redirVal = "../oferts.asp?cmd=oferts&page=" & request("page") & errMsg 
ElseIf Request("redir") = "history" then
	If AddErr = "" Then
		If myApp.GetAfterCartAdd = "Y" Then
			redirVal = "../cart.asp?cmd=cart" & errMsg
			ClosePop = True
		Else
			redirVal = "../oferts.asp?cmd=oferts" & errMsg
			ClosePop = True
		End If
	Else
		Response.Redirect "../flowAlert.asp?DocFlowErr=" & AddErr & "&Item=" & saveHTMLDecode(Request("Item"), False) & "&retURL=" & retURL
	End If
Else
	redirVal = "../cart.asp?cmd=cart&update=Y" & errMsg
End If
If doGoRec and errMsg = "" and AddErr = "" Then redirVal = redirVal & "&loadRec=" & Request("Item") & "&Qty=" & Request("T1") & _
	"&Price=" & Request("precio") & "&SaleType=" & Request("SaleType") & "&ItmEntry=" & ItmEntry & "&RecType=" & RecType
If Request("isFlow") = "Y" Then ClosePop = True %>
<body>
<p><br>
</p>
<!--#include file="../linkForm.asp"-->
<script language="javascript" src="../general.js"></script>
<SCRIPT LANGUAGE="JavaScript">
<% If ClosePop Then %>
	opener.location.href = '<%=redirVal%>';
	window.close();
<% Else %>
doMyLink('<%=Split(redirVal, "?")(0)%>', '<%=Split(Replace(redirVal, "'", "\'"), "?")(1)%>', '');
<% End If %>
</script>
</body>
<!--#include file="../itemFunctions.asp" -->

<%@ Language=VBScript %>
<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="../chkLogin.asp" -->
<!--#include file="../lcidReturn.inc" -->
<!--#include file="../myHTMLEncode.asp" -->
<% 
Dim Curr
%> <html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title></title>
</head>
<%
If userType = "C" Then MainDoc = "default.asp" Else MainDoc = "agent.asp"

AfterCartAdd = myApp.GetAfterCartAdd = "Y"
EnableItemRec = myApp.EnableItemRec
LogPur = myApp.EnableCItemPurLog and userType = "C"
set rv = Server.CreateObject("ADODB.recordset")
set rd = Server.CreateObject("ADODB.recordset")

'Revisa si hay suficiente Stock para agregarlo al shopping cart
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKCheckInvBefAdd" & Session("ID")
cmd.Parameters.Refresh()
cmd("@LogNum") = Session("RetVal")
cmd("@ItemCode") = Request("Item")
cmd("@WhsCode") = Request("whscode")
cmd("@SaleType") = Request("SaleType")
cmd("@FirstQuantity") = CDbl(getNumericOut(Request("addQty")))
If Request("chkAddAll") = "Y" Then cmd("@ChkAddAll") = "Y"
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
		
If Verfy Then
	If Request("linememo") <> "" Then LineMemo = "N'" & Request("linememo") & "'" Else LineMemo = "NULL"
	If Request("ManPrc") <> "" Then ManPrc = Request("ManPrc") Else ManPrc = "N"
	
	ItemHasRec = False
	
	If EnableItemRec Then
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKGetRecItems" & Session("ID")
		cmd("@ItemCode") = Request("Item")
		cmd("@LogNum") = Session("RetVal")
		cmd("@CardCode") = Session("username")
		cmd("@SlpCode") = Session("vendid")
		cmd("@WhsCode") = Request.Form("WhsCode")
		cmd("@branch") = Session("branch")
		cmd("@LanID") = Session("LanID")
		set rs = Server.CreateObject("ADODB.RecordSet")
		conn.execute("set ROWCOUNT 1")
		set rs = cmd.execute()
		ItemHasRec = Not rs.Eof
	End If	
	
	doGoRec = ItemHasRec or TreeType = "S"
	
	If Not doGoRec Then
		If Not VerfyInCart or myApp.BasketMItems Then
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.CommandText = "DBOLKCartAddSF" & Session("ID")
			cmd.CommandType = &H0004
			cmd.ActiveConnection = connCommon
			cmd.Parameters.Refresh
			cmd("@PriceList") = Session("PriceList")
			cmd("@FirstQuantity") = CDbl(getNumericOut(Request("addQty")))
			cmd("@FirstPrice") = CDbl(Replace(getNumericOut(Request("precio")), Request("Currency"), ""))
			cmd("@LogNum") = Session("RetVal")
			cmd("@ItemCode") = Request("item")
			cmd("@WhsCode") = Request("WhsCode")
			cmd("@SaleType") = Request("SaleType")
			cmd("@ManPrc") = ManPrc
			If Request("TaxCode") <> "" Then cmd("@TaxCode") = Request("TaxCode")
			If Request("chkAddAll") = "Y" Then cmd("@All") = "Y"
			cmd("@SlpCode") = Session("vendid")
			cmd("@branch") = Session("branch")
			If Request("Currency") <> "" Then cmd("@ItemCur") = Request("Currency")
			cmd.execute()
			lineNum = cmd("@LineNum")
		'si ya fue agregado se le suma la cantidad que quiere comprar a la que ya fue agregada.
		Else
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKCartUpdateQty" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@LogNum") = Session("RetVal")
			cmd("@LineNum") = lineNum
			cmd("@SaleType") = Request("SaleType")
			cmd("@Quantity") = CDbl(getNumericOut(Request("addQty")))
			cmd.execute()
		End If
		
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKDocLineSaveNote"
		cmd.Parameters.Refresh()
		cmd("@LogNum") = Session("RetVal")
		cmd("@LineNum") = lineNum
		If Request("linememo") <> "" Then cmd("@Note") = Request("linememo")
		cmd.execute()

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
'Esta funcion revisa si el articulo se queda en el wish list o no
If Request("Wish") = "y" Then
	If Request("Wishl") <> "Y" Then
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKWLDel"
		cmd.Parameters.Refresh()
		cmd("@CardCode") = Session("UserName")
		cmd("@ItemCode") = Request("Item")
		cmd.execute()
	End If
End If

Conn.Close
If Request("Wish") = "y" Then %>
	<SCRIPT LANGUAGE="JavaScript">
	opener.location.href = '../<%=MainDoc%>?cmd=wish';
	</SCRIPT>
<% 
End if

If Not doGoRec Then
	If AfterCartAdd Then
		onLoad = "opener.location.href='../cart.asp?focus=frmSmallSearch.string';"
	Else
		onLoad = "opener.loadMinRep();"
	End If
Else
	onLoad = "document.frm.submit();"
End If %>
<body onLoad="<%=onLoad%>">
<script language="javascript">
<% If Not doGoRec  Then

		If myApp.GetAfterCartAdd = "Y" Then
			redirVal = "../cart.asp?cmd=cart" & errMsg
		Else
			redirVal = "../item.asp?item=" & Server.URLEncode(Request("Item")) & "&cmd=a" & errMsg
		End If
		
		Response.Redirect redirVal
End If %>
</script>
<% If doGoRec Then
lineID = ItmEntry & ItmEntry
 %>
<form name="frm" method="post" action="addCartRec.asp">
<input type="hidden" name="Item" value="N'<%=Replace(Request("Item"), "'", "''")%>'">
<input type="hidden" name="Qty<%=lineID%>" value="<%=Request("addQty")%>">
<input type="hidden" name="price<%=lineID%>" value="<%=Request("precio")%>">
<input type="hidden" name="selUn<%=lineID%>" value="<%=Request("SaleType")%>">
<input type="hidden" name="memo<%=lineID%>" value="<%=myHTMLEncode(Request("linememo"))%>">
<input type="hidden" name="whs<%=lineID%>" value="<%=Request("WhsCode")%>">
<input type="hidden" name="ManPrc<%=lineID%>" value="<%=ManPrc%>">
<input type="hidden" name="RecType<%=lineID%>" value="<% If EnableItemRec and ItemHasRec Then %>1<% ElseIf TreeType = "S" Then %>2<% End If %>">
<input type="hidden" name="chkAddAll<%=lineID%>" value="<%=Request("chkAddAll")%>">
</form>
<% End If %>
</body>
</html>
<% else
	If userType = "V" Then 
		response.redirect "itemdetails.asp?Item=" & saveHTMLDecode(Request("Item"), False) & "&cmd=a&chkinv=f&qty=" & getNumeric(Request("addQty")) & "&wish=" & Request("wish")
	ElseIf userType = "C" Then 
		response.redirect "../item.asp?Item=" & saveHTMLDecode(Request("Item"), False) & "&cmd=a&chkinv=f&qty=" & getNumeric(Request("addQty")) & "&wish=" & Request("wish")
	End If
End If
 %>
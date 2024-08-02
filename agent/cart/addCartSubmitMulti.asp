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
<!--#include file="../lcidReturn.inc" -->
<!--#include file="../myHTMLEncode.asp" -->
<%
Dim Curr
If userType = "C" Then MainDoc = "default.asp" Else MainDoc = "agent.asp"
errMsg = ""
 
ArrItems = Split(Request("Item"), "{S}")
ArrQtys = Split(Request("T1"), "{S}")

set rs = server.createobject("ADODB.RecordSet")
LogPur = myApp.EnableCItemPurLog and userType = "C"

set cmdAdd = Server.CreateObject("ADODB.Command")
cmdAdd.ActiveConnection = connCommon
cmdAdd.CommandType = &H0004
cmdAdd.CommandText = "DBOLKCartAddSFM" & Session("ID")
cmdAdd.Parameters.Refresh()
cmdAdd("@lognum") = Session("RetVal")
cmdAdd("@PriceList") = Session("PriceList")
cmdAdd("@UserType") = userType
cmdAdd("@branchIndex") = Session("branch")
cmdAdd("@SlpCode") = Session("SlpCode")

For i = 0 to UBound(ArrItems)
    item = ArrItems(i)
    ItemCode = Mid(item, 3, Len(item)-3)
    qtyVal = CDbl(getNumericOut(ArrQtys(i)))
    
    
	'Revisa si hay suficiente Stock para agregarlo al shopping cart
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKCheckInvBefAdd" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@LogNum") = Session("RetVal")
	cmd("@ItemCode") = ItemCode
	cmd("@FirstQuantity") = qtyVal
	cmd("@branch") = Session("branch")
	cmd("@SlpCode") = Session("vendid")
	cmd("@UserType") = userType
	set rs = cmd.execute()
	
	'Verfy = rv("Verfy") = "Y"
	'ItmEntry = rv("DocEntry")
	'TreeType = rv("TreeType")
	'If TreeType = "S" Then EnableItemRec = False
	'VerfyInCart = rv("VerfyInCart") = "True"
	'lineNum = rv("VerfyInCartLineNum")
	'If TreeType = "S" Then 
	'	EnableItemRec = False
	'	RecType = 2
	'End If
	
	If rs("Verfy") = "Y" Then
        AddItem = True
        If Request("SaleType") <> "" Then SaleType = Request("SaleType") Else SaleType = "NULL"
        If Request("OfertIndex") <> "" Then OfertIndex = Request("OfertIndex") Else OfertIndex = "NULL"
        If Request("precio") <> "" Then Precio = Request("Precio") Else Precio = "NULL"
        If Session("branch") <> "" Then branch = Session("branch") Else branch = "-1"
        TaxCode = ""
        If myApp.LawsSet = "MX" or myApp.LawsSet = "CL" or myApp.LawsSet = "CR" or myApp.LawsSet = "GT" or myApp.LawsSet = "US" or myApp.LawsSet = "CA" Then
      	    If Request("TaxCode") <> "" Then
	      	    TaxCode = Request("TaxCode")
	        Else
	      	    TaxCode = getItemTaxCode(ItemCode)
	        End If
	    
      	    If TaxCode = "Disabled" Then 
      		    TaxCode = "NULL"
      	    'ElseIf TaxCode = "" Then
  			'    errMsg = errMsg & "&errMulti=tax&tItem=" & item
  			'    AddItem = False
  		    End If
        End If
        If AddItem Then
			cmdAdd("@Quantity") = qtyVal
			cmdAdd("@item") = ItemCode
			If Precio <> "NULL" Then cmdAdd("@itemprice") = Precio
			If myApp.LawsSet = "MX" or myApp.LawsSet = "CL" or myApp.LawsSet = "CR" or myApp.LawsSet = "GT" or myApp.LawsSet = "US" or myApp.LawsSet = "CA" Then
				If TaxCode <> "Disabled" and TaxCode <> "" Then cmdAdd("@TaxCode") = TaxCode
			End If
			If OfertIndex <> "NULL" Then cmdAdd("@OfertIndex") = OfertIndex
			cmdAdd.Execute()
			
			
            If Request("redir") = "wish" Then
				set cmd = Server.CreateObject("ADODB.Command")
				cmd.ActiveConnection = connCommon
				cmd.CommandType = &H0004
				cmd.CommandText = "DBOLKWLDel"
				cmd.Parameters.Refresh()
				cmd("@CardCode") = Session("UserName")
				cmd("@ItemCode") = ItemCode
				cmd.execute()
            End If
		      
            If LogPur Then
				set cmd = Server.CreateObject("ADODB.Command")
				cmd.ActiveConnection = connCommon
				cmd.CommandType = &H0004
				cmd.CommandText = "DBOLKAddItemLog" & Session("ID")
				cmd.Parameters.Refresh()
				cmd("@CardCode") = Session("UserName")
				cmd("@ItemCode") = ItemCode
				cmd("@LineType") = "P"
				cmd.execute()
            End If
        End If
    Else
  	    errMsg = errMsg & "&errMInv=" & Mid(item, 2, Len(item)-2)
    End If
Next

Conn.Close
ClosePop = False
If myApp.GetAfterCartAdd = "Y" and Request("redir") <> "history" Then
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
	redirVal = "../default.asp?cmd=clientProm" & errMsg 
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
If Request("isFlow") = "Y" Then ClosePop = True %>
<body>
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

<%@ Language=VBScript %> 
<% If session("ID") = "" Then response.redirect "lock.asp" %>
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

%>
<!--#include file="myHTMLEncode.asp"-->
<!--#include file="lang.asp"-->
<!--#include file="clearItem.asp"-->
<!--#include file="authorizationClass.asp"-->
<!--#include file="topGetValue.asp"-->
<!--#include file="topGetValueSelect.asp"-->

<% 
Dim imgPath
Dim SelMonth, selDay, selYear
Dim tmpStr
Dim monthNames, dayNames
Dim lastDay
Dim tmpLang
dim firstWeekDay
Dim scriptName
Dim myToday ' this day

Dim myAut
set myAut = new clsAuthorization

set rs = Server.CreateObject("ADODB.recordset")

set rd = Server.CreateObject("ADODB.RecordSet")

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKCheckDisLang" & Session("ID")
cmd.Parameters.Refresh()
cmd("@LanID") = Session("LanID")
set rs = cmd.execute()
If Not rs.Eof Then Response.Redirect "operaciones.asp?cmd=home&newLng=" & rs(0)
rs.close
If Session("RetVal") <> "" Then
	If myApp.CopyLastFCRate Then
		cmd.CommandText = "DBOLKCopyLastFCRate" & Session("ID")
		cmd.Parameters.Refresh()
		cmd.execute()
	End If
End If

isWebKit = InStr(LCase(Request.ServerVariables("HTTP_USER_AGENT")), "webkit") <> 0 or InStr(LCase(Request.ServerVariables("HTTP_USER_AGENT")), "mozilla/4") <> 0
%>
<!--#include file="loadAlterNames.asp"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="mobileoptimized" content="0">
<% If isWebKit Then %><meta name="viewport" content="width=320, maximum-scale=1, user-scalable=no"><% End If %>
<title>Mobile OLK</title>
<script language="javascript" src="general.js"></script>
<link rel="stylesheet" href="style.css">
</head>
<% 
curCmd = Request("cmd")

bottomSearchCmd = "searchclient"

Select Case curCmd
	Case "activity", "activityContent", "activityUDF", "activityAddress", "activityGeneral"
		bottomSearchCmd = "openActivitiesSearch"
	Case "slistsearch", "searchclient", "searchitem", "searchCart"
		If curCmd <> "searchitem" or curCmd = "searchitem" and Request("btnRep") = "" Then onLoad = "document.search1.string.focus();"
	Case "delivery"
		onLoad = "document.frmOrder.txtOrderNum.focus();"
	Case "deliveryCheck"
		onLoad = "document.frmSearchItm.txtItem.focus();"
	Case "searchDelCheckItem"
		onLoad = "document.frmConfirm.txtSaleUnit.focus();document.frmConfirm.txtSaleUnit.select();"
	Case "invChkInOutCheckSerial"
		onLoad = "focusSerNum();"
	Case "invChkInOutCheck" 
		onLoad = "document.frmSearchItm.txtItem.focus();"
	Case "invChkInOutAddByPack"
		onLoad = "document.frmAddByPack.txtItem.focus();"
	Case "invChkInOut"
		onLoad = "document.frmOrder.txtOrderNum.focus();"
	Case "addcart"
		onLoad = "document.addcart.Quantity.focus();"
		bottomSearchCmd = "slistsearch"
	Case "cart"
		onLoad = "document.frmAddFast.txtFastAdd.focus();"
End Select
 %>
<body topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" <% If onLoad <> "" Then %>onload="<%=onLoad%>"<% End If %><% If Session("rtl") <> "" Then %> dir="rtl"<% End If %>>
<div<% If isWebKit Then %> style="position: absolute; top: 0px; width: 100%;"<% End If %>>
	<div style="width: 26px; position: absolute; top: 0px; left: 0px; z-index: 1;"><a href="operaciones.asp?cmd=home">
	<img src="images/pocket_art_r1_c1.gif" border="0" alt></a></div>
	<div style="text-align: center; background-color: #FDAF2F; z-index: 0;"><img src="images/pocket_art_r1_c2.gif" border="0" alt></div>
	<div style="width: 26px; position: absolute; top: 0px; right: 0px; z-index: 1;"><a href="operaciones.asp?cmd=about">
	<img src="images/pocket_art_r1_c8.gif" border="0" alt></a></div>
</div>
<div<% If isWebKit Then %>  style="overflow: auto; width: 100%; height: 100%; " id="scroller"<% End If %>>
<% If isWebKit Then %> <div style="height: 24px;"></div><% End If %>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" bordercolor="#111111">
    <tr>
      <td colspan="10" bgcolor="#9BC4FF" style="height: 19px"><% 
      Select Case curCmd
      Case "cartAddMulti" %><!--#include file="cart/addCartMulti.asp" -->
      <% Case "addcart" %><!--#include file="cart/addcart.asp" -->
      <% Case "itemVolRep" %><!--#include file="cart/itemVolRep.asp"-->
      <% Case "searchItems" %><!--#include file="C_art/b1.asp" -->
	  <% Case "crcerror" %><!--#include file="crcerror.asp" -->
      <% Case "itemdetails" %><!--#include file="cart/addcart.asp" -->
      <% Case "adSearch" %><!--#include file="adSearch.asp" -->
      <% Case "cartcancel" %><!--#include file="docdel.asp" -->
      <% Case "actCancel" %><!--#include file="actdel.asp" -->
      <% Case "home" %><!--#include file="home.asp" -->
      <% Case "about" %><!--#include file="about.asp" -->
      <% Case "pendientes" %><!--#include file="listpen.asp" -->
      <% Case "searchPend" %><!--#include file="listpenSearch.asp" -->
      <% Case "openActivities" %><!--#include file="openActivities.asp" -->
      <% Case "openActivitiesSearch" %><!--#include file="openActivitiesSearch.asp" -->
      <% Case "activity" %><!--#include file="activity/activityMain.asp"-->
      <% Case "activityContent" %><!--#include file="activity/activityContent.asp"-->
      <% Case "activityUDF" %><!--#include file="activity/activityUDF.asp"-->
      <% Case "activityAddress" %><!--#include file="activity/activityAddress.asp"-->
      <% Case "activityGeneral" %><!--#include file="activity/activityGeneral.asp"-->
      <% Case "goActivities" %><!--#include file="newactgo.asp" -->
      <% Case "datos" %><!--#include file="carddetails.asp" -->
      <% Case "docgo" %><!--#include file="newdocgo.asp" -->
      <% Case "docdel" %><!--#include file="docdel.asp" -->
      <% Case "cart" %><!--#include file="cart/cart.asp" -->
      <% Case "searchclient" %><!--#include file="searchclients.asp" -->
      <% Case "searchclient2" %><!--#include file="searchclientsa.asp" -->
      <% Case "searchresult" %><!--#include file="cresults.asp" -->
      <% Case "msssbo" %><!--#include file="mensaje_sbo/messagenewSBO.asp" -->
      <% Case "mssolk" %><!--#include file="mensaje_olk/messagenew.asp" -->
      <% Case "reportes" %><!--#include file="reportes/reportes.asp" -->
      <% Case "inv" %><!--#include file="inv/invact.asp" -->
      <% Case "slist" %><!--#include file="C_Art/slist.asp" -->
      <% Case "slistsearch"
      SearchItem = True
       %><!--#include file="C_Art/searchitems.asp" -->
      <% Case "slistsearch2" %><!--#include file="C_Art/searchitemsa.asp" -->
      <% Case "searchitem" %><!--#include file="inv/searchitem.asp" -->
      <% Case "buzon" %><!--#include file="mensaje_olk/buzon.asp" -->
      <% Case "cartSubmitConfirm" %><!--#include file="cart/cartSubmitConfirm.asp" -->
      <% Case "cartSubmit" %><!--#include file="cart/cartSubmit.asp" -->
      <% Case "activitySubmitConfirm" %><!--#include file="activity/activitySubmitConfirm.asp" -->
      <% Case "activitySubmit" %><!--#include file="activity/activitySubmit.asp" -->
      <% Case "messageDetail" %><!--#include file="mensaje_olk/messageDetail.asp" -->
      <% Case "sboMessagePost"%><!--#include file="mensaje_sbo/messagepost.asp" -->
      <% Case "olkMessagePost" %><!--#include file="mensaje_olk/messagepost.asp" -->
      <% Case "recountItemSearch" %><!--#include file="inv/itemresults.asp" -->
      <% Case "recountActiveItem" %><!--#include file="inv/activeitem.asp" -->
      <% Case "cartopt" %><!--#include file="cart/cart_extra.asp" -->
      <% Case "cartExp" %><!--#include file="cart/cart_exp.asp" -->
      <% Case "olkRep" %><!--#include file="Reportes/olkItemReport1.asp" -->
      <% Case "ventasRep" %><!--#include file="Reportes/olkItemReport2.asp" -->
      <% Case "bpriceRep" %><!--#include file="Reportes/olkItemReport3.asp" -->
      <% Case "cart_cp" %><!--#include file="cart/cart_cp.asp" -->
      <% Case "viewRepVals" %><!--#include file="viewRepVals.asp" -->
      <% Case "viewRepValsQry", "adSearchValsQry" %><!--#include file="viewRepValsQry.asp" -->
      <% Case "viewRepValsCL", "adSearchValsCL" %><!--#include file="viewRepValsCL.asp" -->
      <% Case "adSearchValsProp" %><!--#include file="adSearchProp.asp" -->
      <% Case "viewRepValsCal", "adSearchValsCal" %><!--#include file="viewRepValsCal.asp" -->
      <% Case "viewRep" %><!--#include file="viewRep.asp" -->
      <% Case "searchCart" %><!--#include file="cart/searchCart.asp" -->
      <% Case "viewImage" %><!--#include file="viewImage.asp" -->
      <% Case "DocFlowErr" %><!--#include file="flowAlert.asp" -->
      <% Case "invChkInOut" %><!--#include file="inv/delSearchOrder.asp"-->
      <% Case "invChkInOutSearch" %><!--#include file="inv/delSearchOrderResult.asp"-->
      <% Case "invChkInOutCheck" %><!--#include file="inv/delOrderCheck.asp"-->
      <% Case "searchInvChkInOutCheckItem" %><!--#include file="inv/delOrderCheckItemSearch.asp"-->
      <% Case "invChkInOutCheckSubmit" %><!--#include file="inv/delOrderCheckSubmit.asp"-->
      <% Case "invChkInOutCheckSerial" %><!--#include file="inv/delOrderCheckSerial.asp"-->
      <% Case "invChkInOutAddByPack" %><!--#include file="inv/delOrderCheckAddByPack.asp"-->
      <% Case "sec" %><!--#include file="section.inc"-->
      <% Case "cartEditLine" %><!--#include file="cart/cartEditLine.asp"-->
      <% Case "UDFCal" %><!--#include file="UDFCal.asp"-->
      <% Case "UDFQry" %><!--#include file="UDFQry.asp"-->
      <% Case "noaccess" %><!--#include file="noaccess.asp"-->
      <% Case "pendClients" %><!--#include file="crdpen.asp"-->
      <% Case "searchClientPend" %><!--#include file="crdpenSearch.asp"-->
      <% Case "newClient" %><!--#include file="client/addCard.asp"-->
      <% Case "newClientAddData" %><!--#include file="client/addCardAddData.asp"-->
      <% Case "newClientUDF" %><!--#include file="client/addCardUDF.asp"-->
      <% Case "newClientAddresses" %><!--#include file="client/clientAddresses.asp"-->
      <% Case "newClientAddress" %><!--#include file="client/clientAddress.asp"-->
      <% Case "newClientContacts" %><!--#include file="client/clientContacts.asp"-->
      <% Case "newClientContact" %><!--#include file="client/clientContact.asp"-->
      <% Case "clientcancel" %><!--#include file="client/clientdel.asp"-->
      <% Case "newClientSubmit" %><!--#include file="client/clientsubmit.asp"-->
      <% Case "newClientSubmitConfirm" %><!--#include file="client/clientsubmitconfirm.asp"-->
      <% Case "cartBreakDown" %><!--#include file="cartBreakDown.asp"-->
      <% end select %></td>
    </tr>
  </table>
<% If isWebKit Then %> <div style="height: 24px;"></div><% End If %>
</div>
<% If isWebKit Then %> 
<div style="position:absolute; bottom: 0px; width: 100%; height:24px; background-color: #FDAF2F; text-align:center;">
<% 
	Select Case curCmd
			Case "viewRepVals"
				retUrl = "?cmd=reportes"
			Case "purchaseCheckSubmit"
				retUrl = "?cmd=purchaseCheck&txtOrderNum=" & Request("txtOrderNum")
			Case "deliveryCheckSubmit"
				retUrl = "?cmd=deliveryCheck&txtOrderNum=" & Request("txtOrderNum")
			Case "invChkInOutCheckSubmit"
				If Status <> "S" Then 
					retUrl = "?cmd=invChkInOutCheck&txtOrderNum=" & Request("txtOrderNum")
				Else
					retUrl = "javascript:history.go(-1);"
				End If
			Case "adSearch"
				Select Case CInt(Request("adObjID"))
					Case 2
						If MultiAdSearch Then retUrl = "?cmd=searchclient2" Else retUrl = "?cmd=searchclient"
					Case 4
						If MultiAdSearch Then retUrl = "?cmd=slistsearch2&slist=N" Else retUrl = "?cmd=slistsearch&slist=N"
				End Select
			Case Else
				retUrl = "javascript:history.go(-1);"
		End Select %>
		<div style="position:absolute; left: 0px; bottom: 0px;">
		<div style="float: left; width: 25px; z-index: 1;"><a href="<%=retUrl%>">
	  	<img src="images/go_<% If Session("rtl") = "" Then %>back<% Else %>next<% End If %>.jpg" border="0" width="25" height="24"></a></div>
	  	<div style="float: left; width: 25px; z-index: 1;"><a href="javascript:history.go(+1);">
	  	<img src="images/go_<% If Session("rtl") = "" Then %>next<% Else %>back<% End If %>.jpg" border="0" width="25" height="24"></a></div>
	  	</div>
	  	<% If Session("RetVal") = "" Then %>&nbsp;<% Else %><!--#include file="cart/cartMainInfo.asp"--><% End If %>
	  	<div style="position: absolute; bottom: 0px;right : 0px;">
	  	<div style="float: right; width: 25px; z-index: 1;"><a href="default.asp?logout=Y">
	  	<img src="images/lock.jpg" border="0"></a></div>
		<% If myApp.EnableORDR or myApp.EnableOQUT Then %>
	  	<div style="float: right; width: 25px; z-index: 1;"><% If Session("RetVal") <> "" then %><a href="operaciones.asp?cmd=cart"><% end if %>
		<img name="pocket_art_r3_c5" src="images/pocket_art<% If Session("RetVal") = "" then response.write "rollover"%>_r3_c5.gif" border="0" alt width="25" height="24"><% If Session("RetVal") <> "" then %></a><% end if %></div><% End If %>
	  	<% If myAut.HasAuthorization(23) or myAut.HasAuthorization(75) Then %>
	  	<div style="float: right; width: 25px; z-index: 1;">
	  	<a href="operaciones.asp?cmd=<%=bottomSearchCmd%>">
		<img name="pocket_art_r3_c3" src="images/pocket_art_r3_c3.gif" border="0" alt width="25" height="24"></a></div><% End If %>
	  	</div>
</div>
<% Else %>
<table cellpadding="0" bgcolor="#FDAF2F" cellspacing="0" border="0" width="100%">
<tr>
	<% Select Case curCmd
			Case "viewRepVals"
				retUrl = "?cmd=reportes"
			Case "purchaseCheckSubmit"
				retUrl = "?cmd=purchaseCheck&txtOrderNum=" & Request("txtOrderNum")
			Case "deliveryCheckSubmit"
				retUrl = "?cmd=deliveryCheck&txtOrderNum=" & Request("txtOrderNum")
			Case "invChkInOutCheckSubmit"
				If Status <> "S" Then 
					retUrl = "?cmd=invChkInOutCheck&txtOrderNum=" & Request("txtOrderNum")
				Else
					retUrl = "javascript:history.go(-1);"
				End If
			Case "adSearch"
				Select Case CInt(Request("adObjID"))
					Case 2
						If MultiAdSearch Then retUrl = "?cmd=searchclient2" Else retUrl = "?cmd=searchclient"
					Case 4
						If MultiAdSearch Then retUrl = "?cmd=slistsearch2&slist=N" Else retUrl = "?cmd=slistsearch&slist=N"
				End Select
			Case Else
				retUrl = "javascript:history.go(-1);"
		End Select %>
	<td style="width: 25px"><a href="<%=retUrl%>">
  	<img src="images/go_<% If Session("rtl") = "" Then %>back<% Else %>next<% End If %>.jpg" border="0" width="25" height="24"></a></td>
  	<td style="width: 25px"><a href="javascript:history.go(+1);">
  	<img src="images/go_<% If Session("rtl") = "" Then %>next<% Else %>back<% End If %>.jpg" border="0" width="25" height="24"></a></td>
  	<td>&nbsp;</td>
  	<% If myAut.HasAuthorization(23) or myAut.HasAuthorization(75) Then %>
  	<td style="width: 25px">
  	<a href="operaciones.asp?cmd=<%=bottomSearchCmd%>">
	<img name="pocket_art_r3_c3" src="images/pocket_art_r3_c3.gif" border="0" alt width="25" height="24"></a></td><% End If %>
	<% If myApp.EnableORDR or myApp.EnableOQUT Then %>
  	<td style="width: 25px"><% If Session("RetVal") <> "" then %><a href="operaciones.asp?cmd=cart"><% end if %>
	<img name="pocket_art_r3_c5" src="images/pocket_art<% If Session("RetVal") = "" then response.write "rollover"%>_r3_c5.gif" border="0" alt width="25" height="24"><% If Session("RetVal") <> "" then %></a><% end if %></td><% End If %>
  	<td style="width: 25px"><a href="default.asp?logout=Y">
  	<img src="images/lock.jpg" border="0"></a></td>
</tr>
</table>
<% End If %>
</body>
<% conn.close
set rs = nothing %>
</html>
<% Function GetDocCur(ObjectCode)
	If myApp.SVer >= "6" Then
		GetDocCur = "DocCur"
	Else
		Select Case CInt(ObjectCode)
			Case 23
			GetDocCur = "DocCurr"
			Case 17
			GetDocCur = "DocCurr"
			Case 13
			GetDocCur = "DocCur"
		End Select
	End If
End Function %>
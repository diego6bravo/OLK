<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<!--#include file="chkLogin.asp" -->
<!--#include file="clearItem.asp"-->
<!--#include file="myHTMLEncode.asp"-->
<!--#include file="authorizationClass.asp"-->
<!--#include file="loadAlterNames.asp"-->
<!--#include file="lang/itemDetailsData.js.asp" -->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

Dim myAut
set myAut = New clsAuthorization

%>
var txtInvoice = '<%=txtInv%>';
var txtOrdr = '<%=txtOrdr%>';		
var dbName = '<%=Session("olkdb")%>';
var txtClient = '<%=txtClient%>';
var txtAgent = '<%=txtAgent%>';
var txtType = '<%=getitemDetailsDatajsLngStr("DtxtType")%>';
var lblItemDetailsCode = '<%=txtCode%>';
var lblItemDetailsWhs = '<%=getitemDetailsDatajsLngStr("DtxtWarehouse")%>';
var lblItemDetailsQty = '<%=getitemDetailsDatajsLngStr("DtxtQty")%>'
var lblItemDetailsUnit = '<%=getitemDetailsDatajsLngStr("DtxtUnit")%>';
var lblItemDetailsDisc = '<%=getitemDetailsDatajsLngStr("DtxtDiscount")%>';
var lblItemDetailsPrice = '<%=getitemDetailsDatajsLngStr("DtxtPrice")%>';
var lblItemDetailsTotal = '<%=getitemDetailsDatajsLngStr("DtxtTotal")%>';
var lblItemDetailsNote = '<%=getitemDetailsDatajsLngStr("DtxtNote")%>';
var lblItemDetailsChooseFrom = '<%=getitemDetailsDatajsLngStr("LtxtSelFromLst")%>';
var lblItemDetailsConfirm = '<%=getitemDetailsDatajsLngStr("DtxtConfirm")%>';
var lblItemDetailsCancel = '<%=getitemDetailsDatajsLngStr("DtxtCancel")%>';
var lblItemDetailsClose = '<%=getitemDetailsDatajsLngStr("DtxtClose")%>';
var lblItemDetailsBestPrice = '<%=getitemDetailsDatajsLngStr("LtxtBestPrice")%>';
var lblItemDetailsDate = '<%=getitemDetailsDatajsLngStr("DtxtDate")%>';
var txtSalUn = '<%=getitemDetailsDatajsLngStr("DtxtSalUnit")%>';
var txtPackUn = '<%=getitemDetailsDatajsLngStr("DtxtPackUnit")%>';
var txtSAP = '<%=getitemDetailsDatajsLngStr("DtxtSAP")%>';
var txtAVL = '<%=getitemDetailsDatajsLngStr("DtxtAvl")%>';
var txtOnHand = '<%=getitemDetailsDatajsLngStr("LtxtInv")%>';
var txtOLK  = '<%=getitemDetailsDatajsLngStr("DtxtOLK")%>';
var txtWHS = '<%=getitemDetailsDatajsLngStr("DtxtWarehouse")%>';
var txtTaxCode = '<%=getitemDetailsDatajsLngStr("LtxtTaxCode")%>';
var txtNotApply = '<%=getitemDetailsDatajsLngStr("DtxtNotApply")%>';
var txtDate = '<%=getitemDetailsDatajsLngStr("DtxtDate")%>';
var txtSalMet = '<%=getitemDetailsDatajsLngStr("DtxtSalUnit")%>';
var txtNoData = '<%=getitemDetailsDatajsLngStr("DtxtNoData")%>';
var txtVolDiscAvl = '<%=getitemDetailsDatajsLngStr("LtxtVolDiscount")%>';
var txtFilterInCart = '<%=getitemDetailsDatajsLngStr("LtxtFilterInCart")%>';
var txtValOnLst = '<%=getitemDetailsDatajsLngStr("LtxtValOnLst")%>';
var txtItmInCart = '<%=getitemDetailsDatajsLngStr("LtxtUnitInCart")%>';
var txtSalesHideCompTxt = '<%=getitemDetailsDatajsLngStr("LtxtHasComp")%>';
var txtDisNotes = '<%=getitemDetailsDatajsLngStr("LtxtNoteDis")%>';
var txtCartQty = '<%=getitemDetailsDatajsLngStr("LtxtCartQty")%>';
var txtErrItmInv = '<%=getitemDetailsDatajsLngStr("LtxtInvNotAvl")%>';
var txtComponents = '<%=getitemDetailsDatajsLngStr("LtxtComponents")%>';
var txtDetails = '<%=getitemDetailsDatajsLngStr("LtxtDetails")%>';
var txtInv = '<%=getitemDetailsDatajsLngStr("LtxtInv")%>';
var txtBestPrices = '<%=getitemDetailsDatajsLngStr("LtxtBestPrice")%>';
var txtSaleRep = '<%=getitemDetailsDatajsLngStr("LtxtSalesRep")%>';
var txtOLKCommited = '<%=getitemDetailsDatajsLngStr("LtxtOLKCommited")%>';
var txtSAPCommited = '<%=getitemDetailsDatajsLngStr("LtxtSAPCommited")%>';
var txtInvErrMsg = '<%=getitemDetailsDatajsLngStr("LtxtNoAvlInv")%>';
var txtItemAddOK = '<%=getitemDetailsDatajsLngStr("LtxtItemAddOK")%>';
var txtVolDiscount = '<%=getitemDetailsDatajsLngStr("LtxtVolDiscount")%>';
var txtSellAll = '<%=getitemDetailsDatajsLngStr("LtxtSellAll")%>';
var EnSellAll = <%=JBool(myApp.EnSelAll)%>;
var ItemCmd = 'D';
var LawsSet = '<%=myApp.LawsSet%>';
var itemRepLinkBestPrice = <%=JBool(myAut.HasAuthorization(100))%>;
var itemRepLinkInvRep = <%=JBool(myAut.HasAuthorization(101))%>;
var itemRepLinkLastSale = <%=JBool(myAut.HasAuthorization(102))%>;
var itemRep1LinkORDR = <%=JBool(myAut.HasAuthorization(103))%>;
var itemRep1LinkOBS = <%=JBool(myAut.HasAuthorization(103))%>;
var GetShowSalUn = <%=JBool(myApp.GetShowSalUn)%>;
var GetShowRef = <%=JBool(myApp.GetShowRef)%>;
var EnableCartSum = <%=JBool(myApp.EnableCartSum)%>;

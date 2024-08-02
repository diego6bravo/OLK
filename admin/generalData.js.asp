<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<!--#include file="authorizationClass.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

Dim myAut
set myAut = New clsAuthorization
%>
var RateDec = <%=myApp.RateDec%>;
var SumDec = <%=myApp.SumDec%>;
var PriceDec = <%=myApp.PriceDec%>;
var QtyDec = <%=myApp.QtyDec%>;
var PercentDec = <%=myApp.PercentDec%>;
var MeasureDec = <%=myApp.MeasureDec%>;


var dbID = <%=Session("ID")%>;
var GetFormatSep = '<%=Mid(FormatNumber(1000, 2),2,1)%>';
var GetFormatComma = '<%=Mid(FormatNumber(1000, 2),6,1)%>';

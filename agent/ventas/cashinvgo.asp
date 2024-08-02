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
<!--#include file="../myHTMLEncode.asp"-->
<%
           set rs = Server.CreateObject("ADODB.recordset")
           sql = "select OLKCommon.dbo.DBOLKGetCardPList" & Session("ID") & "(N'" & saveHTMLDecode(Request("cl"), False) & "', '" & userType & "') listnum"
           set rs = conn.execute(sql)
           Session("UserName") = saveHTMLDecode(Request.QueryString("cl"), False)
           Session("RetVal") = Request.QueryString("doc")
           Session("PayRetVal") = Request.QueryString("payDoc")
           Session("PriceList") = RS("listnum")
           Session("cart") = "cart"
           Session("PayCart") = True
           Response.Redirect "../cart.asp"
           conn.close
%>

<!--#include file="../lang.asp" -->
<!--#include file="../conn.asp"-->
<!--#include file="../authorizationClass.asp"-->
<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%
Response.Expires = -1
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<html <% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<% 


If Request("itemSmallRep") <> "Y" Then
	If Request("Excell") <> "Y" Then
		myApp.LoadDBConfigData CInt(Request("dbID"))
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "OLKCheckPDFAccess"
		cmd.Parameters.Refresh()
		cmd("@ID") = Session("id")
		cmd("@Rnd") = Request("myRnd")
		cmd.execute()
		If cmd.Parameters.Item(0).value = 0 Then Response.Redirect "../accessDenied.asp"
			
		Session("UserName") = Request("UserName")
		userType = Request("UserType")
		Session("vendid") = Request("vendid")
		Session("useraccess") = Request("useraccess")
	Else
		response.ContentType="application/vnd.ms-excel"
	End If
End If

Dim myAut
set myAut = New clsAuthorization

set rs = Server.CreateObject("ADODB.RecordSet")

If Request("Excell") <> "Y" and Request("itemSmallRep") <> "Y" Then
%>
<link rel="stylesheet" href="../design/<%=GetSelDes%>/style/stylenuevo.css"><% End If %>
</head>
<!--#include file="../lcidReturn.inc"-->
<body topmargin="0" leftmargin="0" link="#4783C5" vlink="#4783C5">
<% viewRepPDF = True
imgAddPath = "../" %>
<!--#include file="../repVars.inc" -->
<!--#include file="viewReport.asp"-->
</body>
</html>
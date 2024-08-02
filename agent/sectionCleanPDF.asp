<%@ Language=VBScript %>
<!--#include file="clsApplication.asp"-->
<!--#include file="authorizationClass.asp"-->
<!--#include file="clsSession.asp"-->
<!--#include file="myHTMLEncode.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

myApp.LoadDBConfigData CInt(Request("dbID"))
userType = Request("UserType")

Session("UserName") = Request("UserName")
Session("vendid") = Request("vendid")
Session("branch") = Request("branch")
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "OLKCheckPDFAccess"
cmd.Parameters.Refresh()
cmd("@ID") = Session("id")
cmd("@Rnd") = Request("myRnd")
cmd.execute()
If cmd.Parameters.Item(0).value = 0 Then Response.Redirect "accessDenied.asp"



Dim myAut
set myAut = New clsAuthorization

set rs = Server.CreateObject("ADODB.recordset")
Select Case userType 
	Case "C"
		sql = "select SelDes from OLKCommon"
		set rs = conn.execute(sql)
 		SelDes = rs("SelDes") 
 	Case Else
 		SelDes = 0
End Select
imgAddPath = "" %>
<!--#include file="design/section.inc"-->
<% set rs = nothing
conn.close %>
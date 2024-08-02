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
<!--#include file="../authorizationClass.asp"-->
<%
response.buffer = true

Dim myAut
set myAut = New clsAuthorization

set rs = Server.CreateObject("ADODB.recordset")
obj = CInt(Request("obj"))
redirURL = ""

If userType = "V" Then
	If myApp.CopyLastFCRate Then
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKCopyLastFCRate" & Session("ID")
		cmd.Parameters.Refresh()
		cmd.execute()
	End If

	set rs = Server.CreateObject("ADODB.RecordSet")
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKCheckRestoreUDF" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@SysID") = "OINV"
	cmd("@ObsID") = "TDOC"
	set rs = cmd.execute()
	If rs(0) = "Y" Then Response.Redirect "../configErr.asp?errCmd=Doc&obj=" & obj

	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKCheckAgentObjectCreation" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@ObjID") = obj
	cmd("@CardCode") = Session("UserName")
	cmd("@UserAccess") = Session("UserAccess")
	cmd("@SlpCode") = Session("vendid")
	rs.close
	rs.open cmd, , 3, 1

	If rs("AsignedSLP") = "Y" Then redirURL = "../configErr.asp?errCmd=AsignedSLP"
	If redirURL = "" Then
		For each itm in rs.Fields
			if itm = "Y" Then Response.Redirect "../configErr.asp?errCmd=Doc&obj=" & obj
		next
	End If
	
End If

If redirURL = "" Then

	Series = myAut.GetObjectProperty(Request("obj"), "S")
	
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = adCmdStoredProc
	cmd.CommandText = "DBOLKCreateDocument" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@ObjectCode") = obj
	cmd("@CardCode") = Session("UserName")
	cmd("@branch") = Session("branch")
	cmd("@UserType") = userType
	cmd("@UserSign") = Session("vendid")
	If Series <> "NULL" Then cmd("@Series") = Series
	cmd.execute()
	Session("PriceList") = cmd("@PriceList")
	Session("RetVal") = cmd("@LogNum")
	

	conn.close
	Session("cart") = "cart"
	Session("PayCart") = False
	Session("PayRetVal") = -1
	Response.Redirect "../cart.asp"
Else %>
<!--#include file="../linkForm.asp"-->
	<script language="javascript" src="../../AGE/general.js"></script>
	<script language="javascript">
	doMyLink('<%=Split(redirURL, "?")(0)%>', '<%=Split(redirURL, "?")(1)%>', '');
	</script>
<% End If %>
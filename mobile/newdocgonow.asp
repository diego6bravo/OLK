<%@ Language=VBScript %>
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>

<!--#include file="myHTMLEncode.asp"-->
<!--#include file="clearItem.asp"-->
<% If session("OLKDB") = "" Then response.redirect "lock.asp" %>
<!--#include file="authorizationClass.asp"-->
<%

Dim myAut
set myAut = New clsAuthorization

obj = CInt(Request("ObjCode"))

AddErr = getNewDocError()
If CStr(AddErr) <> "" Then
	retURL = ""
	For each itm in Request.Form
		If LCase(itm) <> "item" Then
			If retURL <> "" Then retURL = retURL & "{a}"
			retURL = retURL & itm & "{e}" & Request(itm)
		End If
	Next
	For each itm in Request.QueryString
		If LCase(itm) <> "item" Then
			If retURL <> "" Then retURL = retURL & "{a}"
			retURL = retURL & itm & "{e}" & Request(itm)
		End If
	Next
	Response.Redirect "operaciones.asp?cmd=DocFlowErr&DocFlowErr=" & AddErr & "&retURL=" & CleanItem(retURL)
End If

Series = myAut.GetObjectProperty(obj, "S")

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "DBOLKCreateDocument" & Session("ID")
cmd.Parameters.Refresh()
cmd("@ObjectCode") = obj
cmd("@CardCode") = Request("c1")
cmd("@branch") = Session("branch")
cmd("@UserType") = userType
cmd("@UserSign") = Session("vendid")
If Series <> "NULL" Then cmd("@Series") = Series
cmd.execute()
Session("plist") = cmd("@PriceList")
Session("RetVal") = cmd("@LogNum")
Session("UserName") = saveHTMLDecode(Request("c1"), True)

Response.Redirect "operaciones.asp?cmd=cart"

Function getNewDocError()
	RetVal = ""

	set rFlow = Server.CreateObject("ADODB.RecordSet")
	set rChk = Server.CreateObject("ADODB.RecordSet")
	
	sqlFlow = "declare @ObjectCode int set @ObjectCode = " & obj & " " & _
	"select T0.FlowID, T0.Name, Type, Query  " & _
	"from OLKUAF T0  " & _
	"inner join OLKUAF1 T1 on T1.FlowID = T0.FlowID and T1.SlpCode in (" & Session("vendid") & ",-999) " & _
	"inner join OLKUAF2 T2 on T2.FlowID = T0.FlowID " & _
	"where T2.ObjectCode = @ObjectCode and T0.Active = 'Y' and T0.ExecAt = 'D1' "
	
	If Request("DocConf") <> "" Then sqlFlow = sqlFlow & " and T0.FlowID not in (" & Request("DocConf") & ") "
	
	sqlFlow = sqlFlow & " order by Type, [Order] asc"
	'response.redirect "http://www.topmanage.com.pa/query.asp?query=" & sqlFlow
	
	set rFlow = conn.execute(sqlFlow)
	sqlBase = 	"declare @CardCode nvarchar(15) set @CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "' " & _
				"declare @SlpCode int set @SlpCode = " & Session("VendID") & " " & _
				"declare @dbName nvarchar(100) set @dbName = db_name() " & _
				"declare @branch int set @branch = " & Session("branch") & " "
	
	do while not rFlow.eof
		sql = sqlBase & rFlow("Query")
		'response.write sql
		set rChk = conn.execute(sql)
		If not rChk.eof then
			If Not IsNull(rChk(0)) Then
				If lcase(rChk(0)) = lcase("True") Then
					If RetVal <> "" Then RetVal = RetVal & ", "
					RetVal = RetVal & rFlow("FlowID")
					If rFlow("Type") = 0 Then Exit do
				End If
			End If
		End If
	rFlow.movenext
	loop
	getNewDocError = RetVal
End Function

%>
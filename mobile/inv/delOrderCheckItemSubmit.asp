<%@ Language=VBScript %>
<!--#include file="../myHTMLEncode.asp"-->
<% If session("OLKDB") = "" Then response.redirect "../lock.asp" %>
<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

If Request("btnConfirm") <> "" Then
	'If Request("Repack") = "Y" Then Repack = "Y" Else Repack = "N"
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKAddIOCheck" & Session("ID")
	cmd.ActiveConnection = connCommon
	cmd.Parameters.Refresh
	cmd("@LogNum") = Session("IORetVal")
	cmd("@ObjectCode") = Session("ObjCode")
	cmd("@Type") = Session("Type")
	cmd("@DocNum") = Request("txtOrderNum")
	cmd("@WhsCode") = saveHTMLDecode(Session("bodega"), True)
	cmd("@ItemCode") = saveHTMLDecode(Request("ItemCode"), True)
	If Request("NumIn") <> 1 Then cmd("@Unit") = CDbl(Request("txtUnit")) Else cmd("@Unit") = 0
	cmd("@SBUnit") = CDbl(Request("txtSBUnit"))
	cmd("@PackUnit") = CDbl(Request("txtPackUnit"))
	cmd.execute
	ManSerNum = cmd("@ManSerNum") = "Y"
	conn.close
	If Not ManSerNum Then
		response.redirect "../operaciones.asp?cmd=invChkInOutCheck&txtOrderNum=" & Request("txtOrderNum")
	Else
		response.redirect "../operaciones.asp?cmd=invChkInOutCheckSerial&txtOrderNum=" & Request("txtOrderNum") & "&ItemCode=" & Request("ItemCode")
	End If
ElseIf Request("cmd") = "clear" Then
	sql = "declare @LogNum int set @LogNum = " & Session("IORetVal") & " " & _
			"declare @ItemCode nvarchar(20) set @ItemCode = N'" & Request("ItemCode") & "' " & _
			"declare @WhsCode nvarchar(8) set @WhsCode = N'" & Session("bodega") & "' " & _
			"delete R3_ObsCommon..DOC4 where LogNum = @LogNum and LineNum in (select LineNum from R3_ObsCommon..DOC1 where LogNum = @LogNum and ItemCode = @ItemCode and WhsCode = @WhsCode) " & _
			"delete R3_ObsCommon..DOC1 where LogNum = @LogNum and ItemCode = @ItemCode and WhsCode = @WhsCode"
	conn.execute(sql)
	conn.close
	response.redirect "../operaciones.asp?cmd=searchInvChkInOutCheckItem&txtOrderNum=" & Request("txtOrderNum") & "&txtItem=" & Request("txtItem")
End If

%>
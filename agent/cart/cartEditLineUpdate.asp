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
<!--#include file="../myHTMLEncode.asp" -->
<!--#include file="../lcidReturn.inc" -->
<% 
Dim Curr
%> <html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title></title>
</head>
<%
set rv = Server.CreateObject("ADODB.Recordset")
set rs = Server.CreateObject("ADODB.Recordset")

sql = "select (select SDKID from r3_obscommon..tcif where companydb = N'" & Session("OlkDB") & "') As SDKId, " & _
	 "ItemCode, Quantity, WhsCode, UseBaseUn from r3_obscommon..doc1 where lognum = " & Session("RetVal") & " and " & _
	 "LineNum = " & Request("LineNum")
set rs = conn.execute(Sql)

Select Case userType 
	Case "V" 
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKSaveLineData" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LogNum") = Session("RetVal")
		cmd("@LineNum") = Request("LineNum")
		cmd("@WhsCode") = Request("WhsCode")
		If Request("LineMemo") <> "" Then cmd("@LineMemo") = Request("LineMemo")
		If Request("VatGroup") <> "" Then cmd("@VatGroup") = Request("VatGroup")
		If Request("TaxCode") <> "" Then cmd("@TaxCode") = Request("TaxCode")
		cmd.execute()
		
		If cmd("@Verfy") = "Y" Then
	   		SaveUDF
	   		DoClose
	   	Else
	        DoBack
	   	End If
	Case "C"
		SaveUDF
		DoClose
End Select

Sub SaveUDF
	set rv = Server.CreateObject("ADODB.RecordSet")
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKGetUDFCmdSave" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@TableID") = "INV1"
	cmd("@UserType") = userType
	cmd("@OP") = "O"
	rv.open cmd, , 3, 1
	
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKSaveLineUDFData" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@LogNum") = Session("RetVal")
	cmd("@LineNum") = Request("LineNum")
	
	If Not rv.Eof Then
		do while not rv.eof
			strVal = Request(rv("InsertID"))
			If strVal <> "" Then
				Select Case rv("TypeID") 
					Case "B" 
						cmd("@" & rv("InsertID")) = CDbl(getNumericOut(strVal))
					Case "D" 
						cmd("@" & rv("InsertID")) = SaveCmdDate(strVal)
					Case Else
						cmd("@" & rv("InsertID")) = strVal
				End Select
			End If
		rv.movenext
		loop
		cmd.execute()
	End If
End Sub

 %>
<% Sub DoClose %>
<SCRIPT LANGUAGE="JavaScript">
opener.updLineMoreBtn(<%=Request("LineNum")%>, <% If Request("LineMemo") <> "" Then %>true<% Else %>false<% End If %>);
window.close();
</SCRIPT>
<% End Sub %>
<% Sub DoBack %>
<form name="frm" method="post" action="cartEditLine.asp">
<% For each itm in Request.Form
If itm <> "cmd" Then %>
<input type="hidden" name="<%=itm%>" value="<%=Request.Form(itm)%>">
<% 
End If
Next %>
<input type="hidden" name="cmd" value="e">
</form>
<script language="javascript">document.frm.submit();</script>
<% End Sub %>
<body>
</body>
</html>



<script language="javascript" src="general.js"></script>
<!--#include file="myHTMLEncode.asp"-->
<% If Session("VendId") = "" Then response.redirect "default.asp" %>
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
set rs = server.createobject("ADODB.RecordSet")

col = 0
searchStr = Request("searchStr") & "*"
Select Case Request("Type")
	Case "DocLink"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKSearchBPDoc"
		cmd.Parameters.Refresh()
		cmd("@dbID") = Session("ID")
		cmd("@CardCode") = Session("UserName")
		cmd("@searchStr") = Request("DocNum")
		cmd("@DocType") = Request("DocType")
		set rd = cmd.execute()

	Case "Crd"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKTopGetValueSearch" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@searchStr") = searchStr
		cmd("@Type") = 2
		cmd("@SlpCode") = Session("vendid")
		cmd("@1Row") = "Y"
		set rd = cmd.execute()
	Case "TCrd"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKTopGetValueSearch" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@searchStr") = searchStr
		cmd("@Type") = 33
		cmd("@SlpCode") = Session("vendid")
		cmd("@1Row") = "Y"
		set rd = cmd.execute()
	Case "Emp"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKTopGetValueSearch" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@searchStr") = searchStr
		cmd("@Type") = 32
		cmd("@SlpCode") = Session("vendid")
		cmd("@1Row") = "Y"
		set rd = cmd.execute()
	Case "Grp"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKTopGetValueSearch" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@searchStr") = searchStr
		cmd("@Type") = 3
		cmd("@SlpCode") = Session("vendid")
		cmd("@1Row") = "Y"
		col = 1
		set rd = cmd.execute()
	Case "Cty"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKTopGetValueSearch" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@searchStr") = searchStr
		cmd("@Type") = 4
		cmd("@SlpCode") = Session("vendid")
		cmd("@1Row") = "Y"
		col = 1
		set rd = cmd.execute()
	Case "Itm"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKTopGetValueSearch" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@searchStr") = searchStr
		cmd("@Type") = 6
		cmd("@SlpCode") = Session("vendid")
		cmd("@1Row") = "Y"
		set rd = cmd.execute()
	Case "TItm"
		
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKTopGetValueSearch" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@searchStr") = searchStr
		cmd("@Type") = 7
		cmd("@SlpCode") = Session("vendid")
		set rd = cmd.execute()
	Case "ItmGrp"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKTopGetValueSearch" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@searchStr") = searchStr
		cmd("@Type") = 11
		cmd("@SlpCode") = Session("vendid")
		If Session("UserName") <> "" Then cmd("@CardCode") = Session("UserName")
		cmd("@UserType") = userType
		cmd("@branch") = Session("branch")
		If Session("PriceList") <> "" Then cmd("@PriceList") = Session("PriceList")
		cmd("@1Row") = "Y"
		set rd = cmd.execute()
		col = 1
	Case "ItmFrm"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKTopGetValueSearch" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@searchStr") = searchStr
		cmd("@Type") = 10
		cmd("@SlpCode") = Session("vendid")
		If Session("UserName") <> "" Then cmd("@CardCode") = Session("UserName")
		cmd("@UserType") = userType
		cmd("@branch") = Session("branch")
		If Session("PriceList") <> "" Then cmd("@PriceList") = Session("PriceList")
		cmd("@1Row") = "Y"
		set rd = cmd.execute()
		col = 1
	Case "Prj"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKTopGetValueSearch" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@searchStr") = searchStr
		cmd("@Type") = 5
		cmd("@SlpCode") = Session("vendid")
		cmd("@1Row") = "Y"
		set rd = cmd.execute()
	Case "AcctRejReason"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKTopGetValueSearch" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@searchStr") = Request("searchStr")
		cmd("@Type") = 12
		cmd("@SlpCode") = Session("vendid")
		set rd = cmd.execute()
	Case "Slp"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKTopGetValueSearch" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@searchStr") = searchStr
		cmd("@Type") = 8
		cmd("@SlpCode") = Session("vendid")
		cmd("@1Row") = "Y"
		set rd = cmd.execute()
		col = 1
	Case "Usr"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKTopGetValueSearch" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@searchStr") = searchStr
		cmd("@Type") = 9
		cmd("@SlpCode") = Session("vendid")
		cmd("@1Row") = "Y"
		set rd = cmd.execute()
		col = 1
End Select

Dim RetVal
If not rd.eof Then
	RetVal = Replace(rd(col), "'", "\'")
Else
	RetVal = ""
End If

 %>
 <textarea style="width: 100%;" rows="50">
<script language="javascript">
<% If Request("Type") <> "DocLink" Then %>
parent.setValue(0, '<%=RetVal%>', '<%=Request("Type")%>');
<% ElseIf Request("Type") = "DocLink" Then
	If not rd.eof Then %>
	parent.setDocNum(<%=rd(0)%>,<%=rd(1)%>);
	<% Else %>
	parent.setDocNum(null,null);
	<% End If %>
<% End If %>
</script>
</textarea>
<% conn.close
set rs = nothing
set rd = nothing %>
<% response.expires = 0 %>
<% If Session("VendId") = "" Then response.redirect "default.asp" %>
<!--#include file="myHTMLEncode.asp"-->
<!--#include file="lang/topGetValueSelect.asp" -->
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

set rd = server.createobject("ADODB.RecordSet") %>
<!--#include file="loadAlterNames.asp" -->
<html <% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>>
<%
showCol1 = False
PassDesc = Request("PassDesc") = "Y"

searchStr = Replace(Request("Value"),"*","%")
Select Case Request("Type")
	Case "Crd"
		showCol1 = True
		
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKTopGetValueSearch" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@searchStr") = Request("Value")
		cmd("@Type") = 2
		cmd("@SlpCode") = Session("vendid")
		set rd = cmd.execute()

		myTitle = Replace(gettopGetValueSelectLngStr("LttlSelX"), "{0}", Server.HTMLEncode(txtClient))
		colTitle = txtClient
	Case "TCrd"
		showCol1 = True
		
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKTopGetValueSearch" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@searchStr") = Request("Value")
		cmd("@Type") = 33
		cmd("@SlpCode") = Session("vendid")
		set rd = cmd.execute()
		
		myTitle = Replace(gettopGetValueSelectLngStr("LttlSelX"), "{0}", Server.HTMLEncode(txtClientLead))
		colTitle = txtClientLead
	Case "Emp"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKTopGetValueSearch" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@searchStr") = Request("Value")
		cmd("@Type") = 32
		cmd("@SlpCode") = Session("vendid")
		set rd = cmd.execute()
		
		myTitle = Replace(gettopGetValueSelectLngStr("LttlSelX"), "{0}", gettopGetValueSelectLngStr("DtxtOwner"))
		colTitle = gettopGetValueSelectLngStr("LtxtLastName")
	Case "Grp"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKTopGetValueSearch" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@searchStr") = Request("Value")
		cmd("@Type") = 3
		cmd("@SlpCode") = Session("vendid")
		set rd = cmd.execute()
		myTitle = Replace(gettopGetValueSelectLngStr("LttlSelX"), "{0}", gettopGetValueSelectLngStr("DtxtGroup"))
		colTitle = gettopGetValueSelectLngStr("DtxtGroup")
	Case "Cty"
		showCol1 = True
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKTopGetValueSearch" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@searchStr") = Request("Value")
		cmd("@Type") = 4
		cmd("@SlpCode") = Session("vendid")
		set rd = cmd.execute()
		myTitle = Replace(gettopGetValueSelectLngStr("LttlSelX"), "{0}", gettopGetValueSelectLngStr("DtxtCountry"))
		colTitle = gettopGetValueSelectLngStr("DtxtCountry")
	Case "Prj"
		showCol1 = True
		
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKTopGetValueSearch" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@searchStr") = Request("Value")
		cmd("@Type") = 5
		cmd("@SlpCode") = Session("vendid")
		set rd = cmd.execute()
		
		myTitle = Replace(gettopGetValueSelectLngStr("LttlSelX"), "{0}", gettopGetValueSelectLngStr("LtxtPrj"))
		colTitle = gettopGetValueSelectLngStr("LtxtPrj")
	Case "Itm"
		showCol1 = True
		
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKTopGetValueSearch" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@searchStr") = Request("Value")
		cmd("@Type") = 6
		cmd("@SlpCode") = Session("vendid")
		set rd = cmd.execute()
		
		myTitle = Replace(gettopGetValueSelectLngStr("LttlSelX"), "{0}", gettopGetValueSelectLngStr("DtxtItem"))
		colTitle = gettopGetValueSelectLngStr("DtxtDescription")
	Case "TItm"
		showCol1 = True
		
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKTopGetValueSearch" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@searchStr") = Request("Value")
		cmd("@Type") = 7
		cmd("@SlpCode") = Session("vendid")
		set rd = cmd.execute()
		
		myTitle = Replace(gettopGetValueSelectLngStr("LttlSelX"), "{0}", gettopGetValueSelectLngStr("DtxtItem"))
		colTitle = gettopGetValueSelectLngStr("DtxtDescription")
	Case "Slp"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKTopGetValueSearch" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@searchStr") = Request("Value")
		cmd("@Type") = 8
		cmd("@SlpCode") = Session("vendid")
		set rd = cmd.execute()
		
		myTitle = Replace(gettopGetValueSelectLngStr("LttlSelX"), "{0}", Server.HTMLEncode(txtAgent))
		colTitle = txtAgent
	Case "Usr"
	
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKTopGetValueSearch" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@searchStr") = Request("Value")
		cmd("@Type") = 9
		cmd("@SlpCode") = Session("vendid")
		set rd = cmd.execute()
		
		myTitle = gettopGetValueSelectLngStr("LttlUser")
		colTitle = gettopGetValueSelectLngStr("LttlUser")
	Case "ItmGrp"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKTopGetValueSearch" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@searchStr") = Request("Value")
		cmd("@Type") = 11
		cmd("@SlpCode") = Session("vendid")
		If Session("UserName") <> "" Then cmd("@CardCode") = Session("UserName")
		cmd("@UserType") = userType
		cmd("@branch") = Session("branch")
		If Session("PriceList") <> "" Then cmd("@PriceList") = Session("PriceList")
		set rd = cmd.execute()

		myTitle = Replace(gettopGetValueSelectLngStr("LttlSelX"), "{0}", gettopGetValueSelectLngStr("LtxtItmGrp"))
		colTitle = txtAlterGrp
	Case "ItmFrm"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKTopGetValueSearch" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@searchStr") = Request("Value")
		cmd("@Type") = 10
		cmd("@SlpCode") = Session("vendid")
		If Session("UserName") <> "" Then cmd("@CardCode") = Session("UserName")
		cmd("@UserType") = userType
		cmd("@branch") = Session("branch")
		If Session("PriceList") <> "" Then cmd("@PriceList") = Session("PriceList")
		set rd = cmd.execute()

		myTitle = Replace(gettopGetValueSelectLngStr("LttlSelX"), "{0}", Server.HTMLEncode(txtAlterFrm))
		colTitle = txtAlterFrm
End Select
 %>

<head>
<title><%=myTitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" type="text/css" href="design/0/style/stylePopUp.css">
<script language="javascript" src="general.js"></script>
</head>

<body topmargin="0" leftmargin="00" rightmargin="0" bottommargin="0">
<form name="frm">
<table border="0" cellpadding="0" width="100%" id="table1">
	<tr class="CSpecialTlt">
		<td><%=myTitle%>&nbsp;</td>
	</tr>
	<tr>
		<td>
		<table border="0" width="490" id="table2" cellpadding="0">
			<% if not rd.eof then %>
			<tr class="CSpecialTlt2">
				<% If showCol1 Then %><td><%=gettopGetValueSelectLngStr("DtxtCode")%></td><% End If %>
				<td><%=Server.HTMLEncode(colTitle)%></td>
				<% If rd.Fields.Count >= 3 Then %><td><%
				Select Case Request("Type")
					Case "Crd" %><%=gettopGetValueSelectLngStr("DtxtType")%><%
					Case "Emp" %><%=gettopGetValueSelectLngStr("DtxtName")%><%
					Case Else
						Response.Write rd(2).Name 
				End Select %></td><% End If %>
			</tr>
			<% 
			If Request("Type") = "Grp" or Request("Type") = "Cty" or Request("Type") = "ItmGrp" or Request("Type") = "ItmFrm" or Request("Type") = "Slp" or Request("Type") = "Usr" Then
				myValCol = 1
				myDescCol = 0
			Else
				myValCol = 0
				myDescCol = 1
			End If
			do while not rd.eof
			myVal = Replace(Replace(rd(myValCol), "'", "\'"), """", """""")
			If PassDesc Then 
				myVal = myVal & "{S}"
				If Not IsNull(rd(myDescCol)) Then myVal = myVal & Replace(Replace(rd(myDescCol), "'", "\'"), """", """""")
			End If
			 %>
			<tr class="CSpecialTbl" style="cursor: pointer;" onclick="window.returnValue = '<%=Replace(myHTMLEncode(myVal), """", "\u0022")%>'; window.close()" onmouseover="this.style.backgroundColor='#CDE3FC';" onmouseout="this.style.backgroundColor='';">
				<% If showCol1 Then %><td><%=rd(0)%></td><% End If %>
				<td><% If Not IsNull(rd(1)) Then %><%=rd(1)%><% End If %></td>
				<% If rd.Fields.Count >= 3 Then %><td><%
				Select Case Request("Type")
					Case "Crd"
						Select Case rd("CardType")
							Case "C"
								Response.Write txtClient
							Case "S"
								Response.Write gettopGetValueSelectLngStr("DtxtSupplier")
							Case "L"
								Response.Write gettopGetValueSelectLngStr("DtxtLead")
						End Select
					Case Else
						Response.Write rd(2)
				End Select %></td><% End If %>
			</tr>
			<% rd.movenext
			loop
			else %>
			<tr class="CSpecialTbl">
				<td>
				<p align="center"><%=gettopGetValueSelectLngStr("DtxtNoData")%></td>
			</tr>
			<% End If %>
		</table>
		</td>
	</tr>
	<tr class="CSpecialTbl">
		<td align="center">
		<input type="button" name="btnCancel" value="<%=gettopGetValueSelectLngStr("DtxtCancel")%>" onclick="javascript:window.close();"></td>
	</tr>
</table>
</form>
</body>

</html>

<% conn.close
set rd = nothing %>
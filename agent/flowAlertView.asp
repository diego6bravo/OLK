<!--#include file="chkLogin.asp" -->
<!--#include file="lang/flowAlertView.asp" -->
<!--#include file="myHTMLEncode.asp"-->
<html <% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=getflowAlertViewLngStr("LttlDocFlow")%></title>
<link rel="stylesheet" type="text/css" href="design/0/style/stylePopUp.css">
<STYLE TYPE="TEXT/CSS">
<!--
body{
scrollbar-base-color:#014891;
scrollbar-face-color:#0069D2;
scrollbar-highlight-color:#0069D2;
scrollbar-3dlight-color:#014891;
scrollbar-darkshadow-color:#014891;
scrollbar-Shadow-color:#014891;
scrollbar-arrow-color:#FFFFFF;
scrollbar-track-color:#0068D1;
}
.input		
{

	
	color : #3366CC;
	font-family : Verdana, Arial, Helvetica, sans-serif;
	font-size : 10px;
	background-image: url('menybg.gif');
	background-repeat: repeat-x;
	border: 1px solid #555555
}
-->
</STYLE>
</head>
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
set rs = Server.CreateObject("ADODB.RecordSet")
sql = "select T0.FlowID, Name, Type, LineQuery, NoteBuilder, NoteQuery, NoteText, T1.Note, ExecAt " & _
		"from OLKUAF T0 " & _
		"inner join OLKUAF3 T1 on T1.FlowID = T0.FlowID and LogNum = " & Request("LogNum") & " " & _
		"where Type = 2 and Active = 'Y'  " & _
		"order by Type asc, [Order] asc  "
rs.open sql, conn, 3, 1 %>
<body bgcolor="#F0F8FF" topmargin="0" leftmargin="0" rightmargin="0">
<table border="0" cellpadding="0" width="100%" id="table1">
	<% if not rs.eof then
	do while not rs.eof %>
	<tr>
		<td width="777" class="GeneralTlt">
		&nbsp;<%=rs("Name")%></td>
	</tr>
	<tr>
		<td width="775" valign="top" class="GeneralTbl"><%=BuildNote()%>&nbsp;</td>
	</tr>
	<tr>
		<td width="777" class="GeneralTblBold">
		<%=getflowAlertViewLngStr("DtxtNote")%>&nbsp;<%=rs.bookmark%>: <% If Not IsNull(rs("Note")) Then %><%=rs("Note")%><% End If %>
		</td>
	</tr>
	<tr>
		<td width="775" class="GeneralTlt"><span style="font-size: 1pt">&nbsp;</span></td>
	</tr>
	<tr>
		<td width="777" class="GeneralTlt">
		<span style="font-size: 1pt">&nbsp;</span></td>
	</tr>
	<tr>
		<td width="777"><span style="font-size: 1pt">&nbsp;</span></td>
	</tr>
	<% If Not IsNull(rs("LineQuery")) Then %>
	<tr>
		<td width="777">
		<p align="center">
		<iframe name="content" width="100%" src="flowAlertDetails.asp?FlowID=<%=rs("FlowID")%>&LogNum=<%=Request("LogNum")%>&SelDes=0&ExecAt=<%=rs("ExecAt")%>" border="0" frameborder="0" style="border: 3px solid #D6EAFE" height="103">
		Your browser does not support inline frames or is currently configured not to display inline frames.
		</iframe></td>
	</tr>
	<% End If
	rs.movenext
	loop
	else %>
	<tr>
		<td width="777" align="center" class="GeneralTbl">
		<b><%=getflowAlertViewLngStr("DtxtNoData")%></b></td>
	</tr>
	<% End If %>
	<tr>
		<td width="777" class="GeneralTbl">
		<p align="center">
		<input type="button" value="<%=getflowAlertViewLngStr("DtxtClose")%>" name="B2" style="font-family: Verdana; font-size: 10px; color: #FFFFFF; border: 1px solid #000000; background-color: #3366CC" onclick="javascript:window.close()"></td>
	</tr>
</table>
</body>

</html>
<% Function BuildNote()
	myNote = rs("NoteText")
	If rs("NoteBuilder") = "Y" Then
		sqlBase = 	"declare @LogNum int set @LogNum = " & Request("LogNum") & " " & _
					"declare @LanID int set @LanID = " & Session("LanID") & " " & _
					"declare @CardCode nvarchar(15) set @CardCode = (select CardCode from R3_ObsCommon..TDOC where LogNum = @LogNum) " & _
					"declare @SlpCode int set @SlpCode = " & Session("VendID") & " " & _
					"declare @dbName nvarchar(100) set @dbName = db_name() "
		sql = sqlBase & rs("NoteQuery")
		set rNote = Server.CreateObject("ADODB.RecordSet")
		sql = QueryFunctions(sql)
		set rNote = conn.execute(sql)
		If Not rNote.Eof Then
			For each item in rNote.Fields
				If Not IsNull(item) Then
					myNote = Replace(myNote,"{" & item.Name & "}", item)
				End If
			next
		End IF
	End If
	set rNote = nothing
	BuildNote = myNote
End Function %>
<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="../chkLogin.asp" -->
<html <% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>>

<head>
<% 
set rs = Server.CreateObject("ADODB.recordset")
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetObjImg" & Session("ID")
cmd.Parameters.Refresh()
If Request("Item") <> "" Then
	cmd("@ObjectCode") = 4
	cmd("@Code") = Request("Item")
ElseIf Request("card") <> "" Then
	cmd("@ObjectCode") = 2
	cmd("@Code") = Request("card")
End If
set rs = cmd.execute()
varx = rs(0)
SelDes = GetSelDes
           %>
<title><%=Request("Item")%></title>
</head>
<script language="javascript">
function chkMax()
{
	doResize = false;
	MaxW = 541; MaxH = 541;
	NewH = document.body.clientHeight
	NewW = document.body.clientWidth;
	if (NewH < MaxH) { NewH = MaxH; doResize = true; }
	if (NewW < MaxW) { NewW = MaxW; doResize = true; }
	if (doResize) 
	{
		try { window.resizeTo(NewW, NewH); }
		catch (e) { }
	}
}
</script>
<frameset framespacing="0" border="0" frameborder="0" rows="*,120" onresize="chkMax()";>
	<frame name="header" noresize src="ImageMain.asp?pic=<%=varx%>&SelDes=<%=SelDes%>">
	<frame name="main" src="ImageList.asp?SelDes=<%=SelDes%>&item=<%=myQSEncode(Request("item"))%>&card=<%=myQSEncode(Request("card"))%>">
	<noframes>
	<body>

	<p>This page uses frames, but your browser doesn&#39;t support them.</p>

	</body>
	</noframes>
</frameset>
</html>
<% set rs = nothing
conn.close %>
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
<title>Imagenes para el articulo <%=Request("Item")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<script language="javascript">
function changepic(img_src) {
window.mainImage.location.href="ImageMain.asp?pic="+img_src;
}
function setpic(img_src) {
opener.changepic(img_src)
window.close()
}
</script>
</head>
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
           %>
<body topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0">

<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table2" height="344">
	<tr>
		<td>
		<p align="left">
		<iframe name="mainImage" src="ImageMain.asp?pic=<%=varx%>" marginwidth="1" marginheight="1" title="mainImage" align="center" border="0" frameborder="0" height="344" width="620" style="text-align: center">
		Your browser does not support inline frames or is currently configured not to display inline frames.
		</iframe></td>
	</tr>
</table>

<table border="0" cellpadding="0" id="table1">
	<tr>
	<td width="100">
	<iframe name="imageList" src="ImageList.asp?item=<%=Request("item")%>&card=<%=Request("card")%>" width="620" height="132" marginwidth="1" marginheight="1" border="0" frameborder="0">
	Your browser does not support inline frames or is currently configured not to display inline frames.
	</iframe></td>
	</tr>
</table>

</body>
<% set rs = nothing
conn.close %>
</html>
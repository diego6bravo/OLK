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
<html>
<body>
<form method="post" action="../search.asp" name="frmGoWL">
<input type="hidden" name="cmd" value="wish">
<input type="hidden" name="document" value="C">
<input type="hidden" name="orden1" value="<% If myApp.GetDefCatOrdr = "C" Then %>OITM.ItemCode<% Else %>ItemName<% End If %>">
<input type="hidden" name="orden2" value="asc">
<input type="hidden" name="chkWL" value="Y">
<%
		For each itm in Request.Form
			If itm <> "cmd" and itm <> "document" and itm <> "orden1" and itm <> "orden2" Then %>
			<input type="hidden" name="<%=itm%>" value="<%=Request(itm)%>">
<%			End If
		Next
		For each itm in Request.QueryString
			If itm <> "cmd" and itm <> "document" and itm <> "orden1" and itm <> "orden2" Then %>
			<input type="hidden" name="<%=itm%>" value="<%=Request(itm)%>">
<%			End If
		Next 
%>
</form>
<script language="javascript">
document.frmGoWL.submit();
</script>
</body>
</html>
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

<!--#include file="../myHTMLEncode.asp"-->
<script language="javascript" src="../general.js"></script>
<%
set rs = Server.CreateObject("ADODB.RecordSet")
sql = "select IsNull(T0.Street, '') + '<br>' + IsNull(T0.City, '') + ' ' + IsNull(T0.State, '') + ' '  + IsNull(T0.ZipCode, '') + '<br>' +  " & _
"IsNull((select OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRY', 'Name', T0.Country, Name) from ocry where Code = T0.Country), '') Address  " & _
"from CRD1 T0 where CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "' and AdresType = '" & Request("AdresType") & "' and Address = N'" & saveHTMLDecode(Request("Address"), False) & "' "
set rs = conn.execute(sql)
 %>
<body>
<script language="javascript">
	parent.setAddress('<%=Replace(rs(0), "'", "\'")%>', '<%=Request("AdresType")%>');
</script>
</body>
<% conn.close %>
</html>
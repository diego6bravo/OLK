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
<!--#include file="../myHTMLEncode.asp"-->
<!--#include file="lang/go.asp" -->
<html>
<head>
</head>
<body>
<%
set rs = Server.CreateObject("ADODB.recordset")
sql = "select OLKCommon.dbo.DBOLKGetCardPList" & Session("ID") & "(N'" & saveHTMLDecode(Request("cl"), False) & "', '" & userType & "') listnum, PriceList from R3_ObsCommon..TLOGControl where LogNum = " &  Request("doc")
set rs = conn.execute(sql)
If rs("ListNum") = -1 Then
	listNumAlrt = "" & getgoLngStr("LtxtValCLastPurListNu") & ""
ElseIf rs("ListNum") = -2 Then
	listNumAlrt = "" & getgoLngStr("LtxtValCLastDetListNu") & ""
ElseIf IsNull(rs("PriceList")) THen
	sql = "update R3_ObsCommon..TLOGControl set PriceList = " & rs("ListNum") & " where LogNum = " &  Request("doc") & " and PriceList is null"
	conn.execute(sql)
End If

If listNumAlrt <> "" Then %>
<script language="javascript" src="../general.js"></script>
<script type="text/javascript">
<!--
alert('<%=listNumAlrt%>');
history.go(-1);
//-->
</script>
<% Else
	Session("UserName") = saveHTMLDecode(Request("cl"), True)
	Session("RetVal") = Request("doc")
	If IsNull(rs("PriceList")) Then
		Session("PriceList") = RS("listnum")
	Else
		Session("PriceList") = rs("PriceList")
	End If
	If Request("status") = "H" Then
		sql = 	"declare @LogNum int set @LogNum = " & Request("doc") & " " & _
				"update r3_obscommon..tlog set status = 'R' where lognum = @LogNum " & _
				"update OLKUAFControl set Status = 'X', ConfirmDate = getdate(), ConfirmUserSign= " & Session("vendid") & " where ExecAt = 'D3' and ObjectEntry = @LogNum and Status in ('O', 'E') "
		conn.execute(sql)
	end if
	Session("cart") = "cart"
	If Request("payDoc") = "" Then
		Session("PayCart") = False
		Session("PayRetVal") = -1
	Else
		Session("PayCart") = True
		Session("PayRetVal") = Request("payDoc")
	End If
	conn.close
	Response.Redirect "../cart.asp?m=Y"
End IF
%>
</body>
</html>
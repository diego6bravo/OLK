<!--#include file="chkLogin.asp"-->
<!--#include file="myHTMLEncode.asp"-->
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>New Page 1</title>
<script language="javascript" src="general.js"></script>
</head>
<%
If Request("Query") <> "" Then
			set rs = Server.CreateObject("ADODB.recordset")
				
	For i = 0 to UBound(myLanIndex)
		myItm = myLanIndex(i)
		If myItm(4) = CStr(Session("LanID")) Then
			sql = "set language " & myItm(5)
			conn.execute(sql)
			Exit For
		End If
	Next
	sql = ""

	Select Case Request("Type")
		Case "ocrdFilter"
			sql = "select '' from OCRD where 1 = 1 and "
	End Select
	
	sql = sql & Request("Query")
	
	Select Case Request("Type")
		Case "ocrdFilter"
	End Select
End If
%>
<body <% If Request("Query") = "" Then %>onload="parent.VerfyQueryVerified();"<% End If %>>
<% 

errMsg = ""

If Request("Query") <> "" Then
	On Error Resume Next
	set rs = conn.execute(sql)
	If Err.Number <> 0 Then
		If Request("Type") <> "newQueryVar" Then
			errMsg = Replace(Replace(Err.Description,"\","\\"),"'","\'")
		Else
			errMsg = Replace(getverfyQueryLngStr("LtxtValidVar"), "{0}", Request("Query"))
		End If %>
	<script language="javascript">alert('<%=errMsg%>')</script>
	<% Else
End If %>
<script language="javascript">
<% If errMsg = "" Then %>
	parent.VerfyQueryVerified();
<% End If %>
</script>
<% End If %>
</body>
<% If Request("Query") <> "" Then
conn.close
set rs = nothing
End If %>
</html>

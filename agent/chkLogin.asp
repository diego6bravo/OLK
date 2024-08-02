<!--#include file="conn.asp"-->
<!--#include file="lang.asp"-->
<%
Dim myLinks
Dim optProm
Dim optWish
Dim optOfert
Dim optNews
Dim optPolls
Dim optBasket
Dim optCat
Dim optMsg
Dim optRep
Dim optCXC
Dim optMyData
Dim optSecIndex

If Request("closePop") <> "Y" Then 
	If Session("ID") = "" and Request.Cookies("OLKAnon") <> "Y" Then
		If Request("pop") <> "Y" Then 
			Response.Redirect Request("AddPath") & "login.asp"
		ElseIf Request("pop") = "Y" Then
			Response.Redirect Request("AddPath") & "chkLogin.asp?closePop=Y&rParent=" & Request("rParent")
		End If
	End If
Else %>
<script language="javascript">
<% If Request("rParent") <> "Y" Then %>
opener.location.href='<%=Request("AddPath")%>login.asp';
<% Else %>
opener.opener.location.href='<%=Request("AddPath")%>login.asp';
opener.close();
<% End If %>
window.close();
</script>
<% End If

Dim repColSum()
Dim LastColorID
Dim DoColorRowBlink
 %>
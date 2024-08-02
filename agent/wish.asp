<!--#include file="clientInc.asp"-->
<% Select Case userType
Case "C" %><!--#include file="clientTop.asp"-->
<% Case "V"%><!--#include file="agentTop.asp"-->
<% 
If Not myAut.HasAuthorization(1) Then Response.Redirect "unauthorized.asp"
End Select %>
<% If Session("UserName") = "-Anon-" or not optWish Then Response.Redirect MainDoc %>
<!--#include file="searchCart.asp"-->
<% Select Case userType
Case "C" %><!--#include file="clientBottom.asp"-->
<% Case "V" %><!--#include file="agentBottom.asp"-->
<% End Select %>
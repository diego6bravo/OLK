<% If Session("Type") <> "ADM" Then
Session.Abandon
If Request("closePop") <> "Y" and Request("parent") <> "Y" Then 
	If Request("pop") <> "Y" Then 
		Response.Redirect Request("AddPath") & "lock.asp"
	ElseIf Request("pop") = "Y" Then
		Response.Redirect Request("AddPath") & "chkLogin.asp?closePop=Y"
	End If
ElseIf Request("parent") = "Y" Then %>
<script language="javascript">
parent.location.href='<%=Request("AddPath")%>lock.asp';
</script>
<% Else %>
<script language="javascript">
opener.location.href='<%=Request("AddPath")%>lock.asp';
window.close();
</script>
<% End If %>
<% End If %>
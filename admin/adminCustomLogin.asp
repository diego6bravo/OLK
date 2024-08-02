<!--#include file="top.asp" -->
<!--#include file="lang/adminCustomLogin.asp" -->
<% 

set rs = Server.CreateObject("ADODB.RecordSet")
myApp.ConnectCommon

If Request.Form.Count > 0 and Request("txtURL") <> "" Then

	myChar = "ABCDEFGHIJKLMNPQRSTUVWXYZ0123456789"
	charCount = Len(myChar) 
	myKey = ""
	
	Randomize
	For i = 1 To 50 ' password length
		myKey = myKey & Mid( myChar, 1 + Int(Rnd * charCount), 1 )
	Next

	sql = 			"declare @dbName nvarchar(100) set @dbName = N'" & Session("olkdb") & "' " & _  
					"declare @URL nvarchar(256) set @URL = N'" & saveHTMLDecode(Request("txtURL"), False) & "' " & _  
					"declare @AccessKey nvarchar(50) set @AccessKey = '" & myKey & "' " & _  
					"If Exists(select '' from OLKDBU where dbName = @dbName and URL = @URL) begin " & _  
					"	select 'E' Confirm " & _  
					"End Else Begin " & _  
					"	select 'S' Confirm " & _  
					"	declare @ID int set @ID = IsNull((select Max(ID)+1 from OLKDBU), 0) " & _  
					"	insert OLKDBU(dbName, ID, URL, AccessKey) " & _  
					"	values(@dbName, @ID, @URL, @AccessKey) " & _  
					"End " 
	set rs = conn.execute(sql)
	If rs("Confirm") = "E" Then %>
		<script language="javascript">alert('<%=getadminCustomLoginLngStr("LtxtURLExists")%>'.replace('{0}', '<%=Request("txtURL")%>'));</script>
	<% 
	End If 
	rs.close 
ElseIf Request("delID") <> "" Then
	sql = "delete OLKDBU where dbName = N'" & Session("olkdb") & "' and ID = " & Request("delID")
	conn.execute(sql)
End If %>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style type="text/css">
.style1 {
				font-weight: bold;
				background-color: #E1F3FD;
}
.style2 {
				background-color: #E1F3FD;
}
.style3 {
				font-family: Verdana;
				font-size: xx-small;
				color: #4783C5;
}
</style>
</head>

<br>
<table border="0" cellpadding="0" width="100%">
				<tr>
								<td bgcolor="#E1F3FD">&nbsp;<b><font face="Verdana" size="1" color="#31659C"><%=getadminCustomLoginLngStr("LttlCustLogin")%></font></b></td>
				</tr>
				<tr>
								<td bgcolor="#F5FBFE">
								<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1">
								</font>
								<font face="Verdana" size="1" color="#4783C5"><%=getadminCustomLoginLngStr("LttlCustLoginNote")%></font></p>
								</td>
				</tr>
				<tr>
								<td>
								<table border="0" cellpadding="0" width="100%" id="table12">
												<tr>
																<td align="center" class="style1">
																<font size="1" face="Verdana" color="#31659C">
																<%=getadminCustomLoginLngStr("LtxtURL")%>
																</font></td>
																<td align="center" class="style1">
																<font face="Verdana" size="1" color="#31659C">
																<%=getadminCustomLoginLngStr("LtxtAccessKey")%></font></td>
																<td align="center" width="16" class="style2">&nbsp;</td>
												</tr>
												<% 
			sql = "select ID, URL, AccessKey from OLKDBU where dbName = N'" & Session("olkdb") & "'"
			rs.open sql, conn, 3, 1
			If Not rs.Eof Then
			do While NOT RS.EOF %>
												<tr bgcolor="#F3FBFE">
																<td valign="top">
																<span class="style3">
																<%=rs("URL")%></span>
																</td>
																<td valign="top">
																<input type="text" readonly size="70" value="<%=rs("AccessKey")%>">
																</td>
																<td valign="middle" width="16">
																<a href="javascript:if(confirm('<%=getadminCustomLoginLngStr("LtxtConfRemURL")%>'.replace('{0}', '<%=Replace(Server.HTMLEncode(Rs("URL")),"'","\'")%>')))window.location.href='adminCustomLogin.asp?delID=<%=rs("ID")%>';">
																<img border="0" src="images/remove.gif" width="16" height="16"></a></td>
												</tr>
												<input type="hidden" name="ID" value='<%=rs("ID")%>'>
												<% RS.MoveNext
				loop
				Else %>
												<tr>
																<td align="center" class="style1" colspan="3">
																<font size="1" face="Verdana" color="#31659C">
																<%=getadminCustomLoginLngStr("DtxtNoData")%></font></td>
												</tr>
												<% End If %>
												<form name="frmAddURL" action="adminCustomLogin.asp" method="post" onsubmit="return valFrmAddURL();">
												<tr bgcolor="#F3FBFE">
																<td valign="top">
																<span class="style3">
																<input type="text" size="50" name="txtURL" value="">&nbsp;<input type="submit" name="btnAdd" value="<%=getadminCustomLoginLngStr("DtxtAdd")%>"></span>
																</td>
																<td valign="top">
																&nbsp;
																</td>
																<td width="16">
																&nbsp;</td>
												</tr>
												</form>
								</table>
								</td>
				</tr>
</table>
<script type="text/javascript">
<!--
function valFrmAddURL()
{
	if (document.frmAddURL.txtURL.value == '')
	{
		alert('<%=getadminCustomLoginLngStr("LtxtEnterURL")%>');
		document.frmAddURL.focus();
		return false;
	}
	return true;
}
//-->
</script>
<!--#include file="bottom.asp" -->
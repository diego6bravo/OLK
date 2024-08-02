<% addLngPathStr = "searchInc/" %>
<!--#include file="lang/searchClients.asp" -->
<table border="0" cellpadding="0" cellspacing="0" width="93%"=>
	<form method="POST" action="clientsSearch.asp" name="frmSmallSearch">
		<tr>
			<td>
			<p align="center"><b>
			<font size="1" face="Verdana" color="#FFFFFF"><%=myHTMLDecode(getsearchClientsLngStr("LtxtSearch"))%>:</font></b></td>
		</tr>
		<tr>
			<td>
			<input class="input" type="text" name="string" size="16" style="border:1px solid #FFFFFF; font-family: Verdana; font-size: 10px; color:#FFFFFF; background-color:#0065CE" value="<%=myHTMLEncode(Request("string"))%>" onfocus="this.select()" accesskey="<% If Session("myLng") = "es" or Session("myLng") = "pt" Then %>B<% ElseIf Session("myLng") = "he" Then %>&#1495;<% Else %>S<% End If %>"></td>
		</tr>
	<tr>
			<td style="font-size: 4px">&nbsp;</td>
		</tr>
		<tr>
			<td>
			<p align="center">
							<input type="submit" value="<%=getsearchClientsLngStr("DbtnSearch")%>" name="B1" style="color: #FFFFFF; font-family: Verdana; font-size: 7pt; border: 1px solid #FFFFFF; background-color: #0066CB; width:76"></td>
		</tr>
		<tr>
			<td style="font-size: 4px">&nbsp;</td>
		</tr>
		<tr>
			<td>
			<p align="center">
			<% ObjID = 2 %>
			<!--#include file="adCustomSearchInc.asp"--></td>
		</tr>
      	<input type="hidden" name="cmd" value="clientsSearch">
		<input type="hidden" name="focus" value="frmSmallSearch.string">
	</form>
</table>

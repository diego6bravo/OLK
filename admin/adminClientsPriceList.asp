<!--#include file="top.asp" -->
<!--#include file="lang/adminClientsPriceList.asp" -->
<% conn.execute("use [" & Session("olkdb") & "]") %>

<br>
<table border="0" cellpadding="0" width="100%" id="table3">
	<tr>
		<td bgcolor="#E1F3FD"><b><font color="#FFFFFF">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><%=getadminClientsPriceListLngStr("LttlPListCDef")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1"> 
		</font><font face="Verdana" size="1" color="#4783C5">
		<%=getadminClientsPriceListLngStr("LttlPListCDefNote")%></font></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE" align="center">
		<br>
		<img border="0" src="images/sboPriceList_200<%=Get200x%>_<%=Session("myLng")%>.jpg">
		<br>
		</td>
	</tr>
	</table>

<%
Function Get200x()
	sql = "select Left(Version, 2) from CINF"
	set rs = conn.execute(sql)
	If CInt(rs(0)) >= 68 Then
		Get200x = "5"
	Else
		Get200x = "4"
	End If
End Function %><!--#include file="bottom.asp" -->
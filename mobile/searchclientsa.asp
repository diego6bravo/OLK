<!--#include file="lang/searchclientsa.asp" -->
<%
   
sql = "select T0.ID, IsNull(T1.AlterName, T0.Name) Name  " & _
		"from OLKCustomSearch T0 " & _
		"left outer join OLKCustomSearchAlterNames T1 on T1.ObjectCode = T0.ObjectCode and T1.ID = T0.ID and T1.LanID = " & Session("LanID") & " " & _
		"where T0.ObjectCode = 2 and T0.Status = 'Y' and exists(select '' from OLKCustomSearchSession where ObjectCode = T0.ObjectCode and ID = T0.ID and SessionID = 'P') " & _
		"order by T0.Ordr "
set rSearch = Server.CreateObject("ADODB.RecordSet")
rSearch.open sql, conn, 3, 1
%>

<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
    <tr>
      <td bgcolor="#9BC4FF">
      <table border="0" cellpadding="0" bordercolor="#111111" width="100%" id="AutoNumber1">
        <tr>
          <td width="100%">
          <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><%=getsearchclientsaLngStr("LtxtAdvClientSearch")%></font></b></td>
        </tr>
        <tr>
          <td width="100%">
			<table border="0" cellpadding="0" width="100%" cellspacing="1">
				<% do while not rSearch.eof %>
				<tr>
					<td width="17">
					<p align="center">
					<img border="0" src="images/icon_menu.gif"></td>
					<td>
					<font face="Verdana" size="1"><a href="javascript:goAdSearch(<%=rSearch("ID")%>);"><%=rSearch("Name")%></a></font></td>
				</tr>
				<% rSearch.movenext
				loop %>
			</table>
          </td>
        </tr>
        </table>
      </td>
    </tr>
    </table>
  </center>
</div>
<script type="text/javascript">
<!--
function goAdSearch(ID)
{
	window.location.href='operaciones.asp?cmd=adSearch&ID=' + ID + '&adObjID=2&slist=<%=Request("slist")%>';
}
//-->
</script>
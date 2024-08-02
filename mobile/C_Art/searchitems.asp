<% addLngPathStr = "C_Art/" %>
<!--#include file="lang/searchitems.asp" -->
<% If Request.Form("plist") <> "" then 
Session("plist") = Request.Form("plist") 
	slist = "Y" 
ElseIf Request("slist") = "Y" Then
	slist = "Y"
Else
	slist = "N" 
End If
	
%>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
    <tr>
      <td bgcolor="#9BC4FF">
      <table border="0" cellpadding="0" bordercolor="#111111" width="100%">
      <!--#include file="CardNameAdd.asp" -->
        <tr>
          <td width="100%">
          <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><%=getsearchitemsLngStr("DtxtCat")%> - <%=getsearchitemsLngStr("LtxtItemsSearch")%>
          </font></b></td>
        </tr>
        <tr>
          <td width="100%">
          <form method="POST" action="operaciones.asp" name="search1">
            <table border="0" cellpadding="0" cellspacing="1"  bordercolor="#111111" width="100%">
              <tr>
                <td width="100%" bgcolor="#66A4FF">
                <p align="center"><b><font size="1" face="Verdana"><%=getsearchitemsLngStr("DtxtSearch")%></font></b></td>
              </tr>
        
             <tr>
                <td width="100%">
                <p align="center">
    <font face="Verdana" size="2" color="#336699"><b>
    <input type="text" name="string" size="25" style="font-size:12px;"></b></font></td>
              </tr>
              <tr>
                <td width="100%">
 <p align="left">
    <font face="Verdana" size="2" color="#336699"><b>
    </b></font></td>
              </tr>
              <% If myApp.SearchExactP Then %>
              <tr>
                <td width="100%">
                <p align="center">
				<font face="Verdana">
				<input type="radio" value="E" name="rdSearchAs" id="rdSearchAsE" <% If myApp.SearchMethodP = "E" Then %>checked<% End If %>><label for="rdSearchAsE"><font size="2"><%=getsearchitemsLngStr("LtxtExact")%></font></label><input type="radio" name="rdSearchAs" id="rdSearchAsS" value="S" <% If myApp.SearchMethodP = "L" Then %>checked<% End If %>><label for="rdSearchAsS"><font size="2"><%=getsearchitemsLngStr("LtxtLike")%></font></label></font></td>
              </tr>
              <% Else %>
              <input type="hidden" name="rdSearchAs" value="S">
              <% End If %>
              <tr>
                <td width="100%">
                <p align="center">
				<input type="submit" value="<%=getsearchitemsLngStr("DbtnSearch")%>" name="B1" style="height: 26px"> 
				<% 
				sql = 	"select T0.ID  " & _
						"from OLKCustomSearch T0 " & _
						"left outer join OLKCustomSearchAlterNames T1 on T1.ObjectCode = T0.ObjectCode and T1.ID = T0.ID and T1.LanID = " & Session("LanID") & " " & _
						"where T0.ObjectCode = 4 and T0.Status = 'Y' and exists(select '' from OLKCustomSearchSession where ObjectCode = T0.ObjectCode and ID = T0.ID and SessionID = 'P') " & _
						"order by T0.Ordr "
				set rSearch = Server.CreateObject("ADODB.RecordSet")
				rSearch.open sql, conn, 3, 1
				If rSearch.recordcount > 0 Then %>
				- <input type="button" value="<%=getsearchitemsLngStr("LtxtAdvanced")%>" name="B2" onclick="javascript:window.location.href='operaciones.asp?cmd=<% If rSearch.recordcount > 1 Then %>slistsearch2<% Else %>adSearch&ID=<%=rSearch("ID")%>&adObjID=4<% End If %>&slist='+document.search1.slist.value;"><% End If %></td>
              </tr>
            </table>
          	<input type="hidden" name="cmd" value="searchItems">
          	<input type="hidden" name="slist" value="<%=slist%>">
          </form>
          </td>
        </tr>
        </table>
      </td>
    </tr>
    </table>
  </center>
</div>
<script type="text/javascript">
function onScan(ev){
var scan = ev.data;
	document.search1.string.value = scan.value;
	document.search1.submit();
}
function onSwipe(ev){
}

try
{
document.addEventListener("BarcodeScanned", onScan, false);
document.addEventListener("MagCardSwiped", onSwipe, false);
}
catch(err) {}
</script>
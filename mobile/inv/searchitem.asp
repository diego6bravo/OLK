<% addLngPathStr = "inv/" %>
<!--#include file="lang/searchitem.asp" -->
<% If Request.Form("Bodega") <> "" Then Session("Bodega") = Request.Form("Bodega") %>
<% If Request("btnSearch") <> "" Then %>
<script type="text/javascript">
<!--

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

//-->
</script>
<form name="search1" method="POST" action="operaciones.asp">
<input type="hidden" name="cmd" value="recountItemSearch">
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
    <tr>
      <td bgcolor="#9BC4FF">
      <table border="0" cellpadding="0" bordercolor="#111111" width="100%" id="AutoNumber1">
        <tr>
          <td width="100%">
          <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><%=getsearchitemLngStr("LtxtInvRecount")%> 
          </font></b></td>
        </tr>
        <tr>
          <td width="100%">
          <p align="center"><font size="1" face="Verdana"><%=getsearchitemLngStr("LtxtItemSearchNote")%></font></td>
        </tr>
        <tr>
          <td width="100%">
            <table border="0" cellpadding="0" cellspacing="1"  bordercolor="#111111" width="100%" id="AutoNumber2">
              <tr>
                <td width="100%">
                <div align="center">
                  <center>
                  <table border="0" cellpadding="0" cellspacing="1" style="border-collapse: collapse" bordercolor="#111111" width="100" id="AutoNumber3">
                    <tr>
                      <td width="100%">
        <input type="text" name="string" size="20"></td>
                    </tr>
                    <tr>
                      <td width="100%">
                      <p align="center">
        <input type="submit" name="btnSearch" value="<%=getsearchitemLngStr("DbtnSearch")%>"></td>
                    </tr>
                  </table>
                  </center>
                </div>           
                </td>
              </tr>
              <tr>
                <td width="100%"><hr color="#408CFF" size="1"></td>
              </tr>
            </table>
          </td>
        </tr>
        </table>
      </td>
    </tr>
    </table>
  </center>
</div>
</form>
<center>
<form action="operaciones.asp" method="post">
	<input type="submit" value="<%=getsearchitemLngStr("LtxtRep")%>" name="btnRep">
	<input type="hidden" name="cmd" value="searchitem">
	</form>
</center>
<% Else %><%
set rs = Server.CreateObject("ADODB.recordset")
sql = "select ItemCode, OnHand, Counted from oitw where wascounted = 'Y' and whscode = '" & Session("Bodega") & "'"
set rs = conn.execute(sql)
%><div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
        <tr>
          <td width="100%" bgcolor="#9BC4FF">
          <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><%=getsearchitemLngStr("LtxtInvRecount")%>
          </font></b></td>
        </tr>
    <tr>
      <td bgcolor="#9BC4FF">
			<table border="0" cellpadding="0" cellspacing="1"  bordercolor="#111111" width="100%" id="AutoNumber4">
                 <tr>
                    <td align="center" colspan="2" bgcolor="#66A4FF"><b>
                    <font size="1" face="Verdana"><%=getsearchitemLngStr("DtxtCode")%></font></b></td>
                    <td align="center" bgcolor="#66A4FF"><b>
                    <font size="1" face="Verdana"><%=getsearchitemLngStr("LtxtInventory")%></font></b></td>
                    <td align="center" bgcolor="#66A4FF"><b>
                    <font size="1" face="Verdana"><%=getsearchitemLngStr("DtxtCounted")%></font></b></td>
                  </tr>
                 <% do While not rs.eof %>
                  <tr>
                    <td width="12"><a href="operaciones.asp?cmd=recountActiveItem&item=<%=Replace(Replace(Replace(rs("ItemCode"),"#","%23"),"&","%26"),"""","%22")%>">
                    <img border="0" src="images/<%=Session("rtl")%>flechaselec.gif"></a></td>
                    <td><font size="1" face="Verdana"><%=RS("ItemCode")%></font></td>
                    <td><font size="1" face="Verdana">
					<p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><%=RS("OnHand")%></font></td>
                    <td><font size="1" face="Verdana">
					<p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><%=RS("Counted")%></font></td>
                  </tr>
                  <% rs.movenext
                  loop %>
			</table>
      </td>
    </tr>
          <form action="operaciones.asp" method="post">
              <tr> 
<td width="100%" bgcolor="#9BC4FF">
                <p align="center">
				<input type="submit" value="<%=getsearchitemLngStr("DbtnSearch")%>" name="btnSearch" style="height: 26px"></td>
              </tr>
            	<input type="hidden" name="cmd" value="searchitem">
          </form>
    </table>
  </center>
</div>
    <% End If %>
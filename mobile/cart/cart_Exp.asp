<% addLngPathStr = "cart/" %>
<!--#include file="lang/cart_Exp.asp" -->
<% set rs = server.createobject("ADODB.RecordSet")
	sql = 	"SELECT T0.ExpnsCode, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OEXD', 'ExpnsName', T0.ExpnsCode, T0.ExpnsName) ExpnsName, IsNull(T1.LineTotal,0) Price " & _
			"FROM OEXD T0 " & _
			"left outer join R3_ObsCommon..DOC3 T1 on T1.ExpnsCode = T0.ExpnsCode and T1.LogNum = " & Session("RetVal")
	set rs = conn.execute(sql)
	%>
<script language="javascript">
function IsNumeric(sText)
{
   var ValidChars = "0123456789<%=GetFormatDec()%>";
   var IsNumber=true;
   var Char;

 
   for (i = 0; i < sText.length && IsNumber == true; i++) 
      { 
      Char = sText.charAt(i); 
      if (ValidChars.indexOf(Char) == -1) 
         {
         IsNumber = false;
         }
      }
   return IsNumber;
   
   }

function FormatNumber(expr, decplaces) 
{
	return formatNumberDec(expr, decplaces, false);
}

function chkNum(Field, Max) {
if (IsNumeric(Field.value)==false) {
	Field.value = FormatNumber(Max,<%=myApp.PriceDec%>)
	alert("<%=getcart_ExpLngStr("DtxtValNumVal")%>") }
	else { Field.value = FormatNumber(Field.value, <%=myApp.PriceDec%>) }
}
</script>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111" id="table3">
    <tr>
      <td bgcolor="#9BC4FF">
      <form method="POST" action="cart/cartupdate2.asp" name="frmCart">
        <table border="0" cellpadding="0" bordercolor="#111111" width="100%" id="table4">
        <tr>
          <td width="100%">
          <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><%=getcart_ExpLngStr("LtxtShopCart")%>&nbsp;-&nbsp;<%=getcart_ExpLngStr("LtxtExpenses")%></font></b></td>
        </tr>
        <tr>
          <td width="100%" style="border-bottom-style: solid; border-bottom-width: 1px">
          <table border="0" cellpadding="0" bordercolor="#111111" width="100%" id="table5">
            <tr>
              <td width="100%">
          <table border="0" cellpadding="0" cellspacing="1"  bordercolor="#111111" width="100%" id="table7">
        <% do while not rs.eof %>
            <tr>
              <td width="11%" height="9" bgcolor="#95BFFF">
              <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
				<input <% If CDbl(rs("Price")) > 0 Then %>checked<% End If %> type="checkbox" name="chkExpns" value="<%=rs("ExpnsCode")%>" id="fp<%=rs("ExpnsCode")%>"></td>
              <td width="52%" height="9" bgcolor="#75ACFF">
              <font size="1" face="Verdana">&nbsp;<b><label for="fp<%=rs("ExpnsCode")%>"><%=rs("ExpnsName")%></label></b></font></td>
              <td width="36%" height="9" bgcolor="#75ACFF" align="right"><b>
              <font size="1" face="Verdana" color="#FF0000">
              <input name="Price<%=rs("ExpnsCode")%>" size="10" style="font-family: Verdana; font-size: 10px; color: #000000; text-align:right; " value="<%=FormatNumber(rs("Price"),myApp.PriceDec)%>" onchange="chkNum(this,<%=rs("Price")%>)" onclick="this.selectionStart=0;this.selectionEnd=this.value.length;"></font></b></td>
            </tr>
         <% rs.movenext
         loop %>
            </table>
              </td>
            </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td width="100%" style="font-size: 5px; border-left-width: 1px; border-right-width: 1px; border-top: 1px solid #FF9933; border-bottom-width: 1px">&nbsp;</td>
        </tr>
        <TR>
        <td>
          <div align="center">
            <center>
            <table border="0" cellpadding="0" cellspacing="1" style="border-collapse: collapse" bordercolor="#111111" width="95" id="table8">
              <tr>
                <td>
                <p align="center"><input type="image" name="I2" border="0" src="images/ok_icon.gif"></td>
                <td>
                <p align="center"><a href="operaciones.asp?cmd=cart"><img border="0" src="images/x_icon.gif"></a></td>
              </tr>
            </table>
            </center>
          </div>
        </td></TR>
    </table>
    	<input type="hidden" name="cmd" value="cartExp">
    </td></tr></table>
    </form>
  </center>
</div>
<% set rs = nothing %>
<% addLngPathStr = "Reportes/" %>
<!--#include file="lang/olkItemReport3.asp" -->
<%
set rw = server.CreateObject("ADODB.recordset")
sql = 	"select T0.Itemcode, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'ItemName', T2.ItemCode, T2.ItemName) ItemName, " & _
		"OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRD', 'CardName', T3.CardCode, T3.CardName) CardName, " & _
		"T0.Currency, Price, Quantity " & _
		"from inv1 T0 " & _
		"inner join oinv T1 on T1.docentry = T0.docentry " & _
		"inner join OITM T2 on T2.ItemCode = T0.ItemCode " & _
		"inner join OCRD T3 on T3.CardCode = T1.CardCode " & _
		"where T1.cardcode = N'" & saveHTMLDecode(Session("UserName"), False) & "' and T0.ItemCode = N'" & saveHTMLDecode(Request("Item"), False) & "' and UseBaseUn = 'N' and Price = " & _
		"(select Min(Price) from inv1 inner join oinv on oinv.docentry = inv1.docentry  " & _
		"where cardcode = T1.CardCode and ItemCode = T0.ItemCode and UseBaseUn = 'N') "
		
		  If Not myAut.HasAuthorization(97) Then sql = sql & " and T1.SlpCode = " & Session("vendid") & " "
		
		sql = sql & "order by T1.DocDate Desc"
    			  rw.open sql, conn,3,1 %>
	<table border="0" cellpadding="0" cellspacing="1"  bordercolor="#111111" width="100%" id="AutoNumber5">
            <% If Not rw.eof then %>
          <tr>
            <td width="100%" colspan="4" height="12">
			<p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><%=getolkItemReport3LngStr("LtxtBestPriceCurClien")%></font></b></td>
            </tr>
          <tr>
            <td width="100%" colspan="4" height="12" bgcolor="#E2F0FE" style="color: #808080"><b>
			<font face="Verdana" size="1"><font color="#000000"><%=getolkItemReport3LngStr("DtxtClient")%> -</font> <%=RW("CardName")%>
			</font></b></td>
            </tr>
          <tr>
            <td width="100%" colspan="4" height="12" bgcolor="#E2F0FE" style="color: #808080"><b>
			<font size="1" face="Verdana"><font color="#000000"><%=getolkItemReport3LngStr("LtxtReference")%></font> <%=RW("Itemcode")%> : <%=RW("Itemname")%></font></b></td>
            </tr>
          <tr>
            <td width="54%" colspan="3" height="12" bgcolor="#D2E9FF"><b>
            <font size="1" face="Verdana"><nobr><%=getolkItemReport3LngStr("LtxtBestPrice")%>&nbsp;<% If Rw.RecordCount > 0 Then %><%=RW("Currency")%>&nbsp;<%=FormatNumber(RW("Price"),myApp.PriceDec)%><% end if %></nobr></font></b></td>
            <td width="46%" height="12" bgcolor="#D2E9FF"><b>
            <font face="Verdana" size="1"><%=getolkItemReport3LngStr("DtxtQty")%>: <% If Rw.RecordCount > 0 Then %><%=RW("Quantity")%><% end if %></font></b></td>
          </tr>
          <tr>
            <td align="center" bgcolor="#D2E9FF">
            <b><font size="1" face="Verdana"><%=txtInv%></font></b></td>
            <td align="center" bgcolor="#D2E9FF">
            <b><font size="1" face="Verdana"><%=getolkItemReport3LngStr("DtxtDate")%></font></b></td>
            <td align="center" bgcolor="#D2E9FF"><b>
            <font size="1" face="Verdana"><%=getolkItemReport3LngStr("DtxtQty")%></font></b></td>
            <td align="center" bgcolor="#D2E9FF"><b>
            <font size="1" face="Verdana"><%=getolkItemReport3LngStr("DtxtPrice")%></font></b></td>
          </tr>
    <% rw.close
    	sql = "select Top 10 DocNum, LineNum+1 As LineNum, Currency, Convert(int,oinv.DocDate), OINV.DocDate, Quantity, Price " & _
    		 "from inv1 inner join oinv on oinv.docentry = inv1.docentry " & _
    		 "where cardcode = N'" & saveHTMLDecode(Session("UserName"), False) & "' and ItemCode = N'" & saveHTMLDecode(Request("Item"), False) & "' and UseBaseUn = 'N' order by 4 Desc"
    rw.open sql, conn,3,1 %>
      	<% do while not rw.eof %>
          <tr>
          <form method="post" action="../cxcdocdetail.asp" target="_blank"> 
            <input type="hidden" name="docnum" value="<%=RW("DocNum")%>"><input type="hidden" name="doctype" value="13">
            <input type="hidden" name="high" value="<%=RW("LineNum")%>">
            <td bgcolor="#E2F0FE"><font size="1" face="Verdana"><%=RW("DocNum")%> (<%=RW("LineNum")%>)</font></td></form>
            <td bgcolor="#E2F0FE">
            <p align="center"><font size="1" face="Verdana"><%=FormatDate(RW("DocDate"), True)%></font></td>
            <td bgcolor="#E2F0FE">
            <p align="center"><font size="1" face="Verdana"><%=RW("Quantity")%></font></td>
            <td bgcolor="#E2F0FE">
            <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><nobr><font size="1" face="Verdana"><%=RW("Currency")%>&nbsp;<%=FormatNumber(RW("Price"),myApp.PriceDec)%></font></nobr></td></tr>
          <% rw.movenext
          loop %>
          <% Else %>
          <tr>
            <td colspan="4" height="12" bgcolor="#D2E9FF" colspan="3"><b>
            <font size="1" face="Verdana"><%=getolkItemReport3LngStr("DtxtNoData")%></font></b></td>
          </tr>
          <% End If %>
                  <tr>
            <td bgcolor="#E2F0FE" colspan="4"><font size="1" face="Verdana">&nbsp;
			</font></td>
            </tr>
          
          </table>
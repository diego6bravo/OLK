<% addLngPathStr = "Reportes/" %>
<!--#include file="lang/olkItemReport1.asp" -->
<%
set rs = Server.CreateObject("ADODB.recordset")
sql = 	"select OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'ItemName', OITM.ItemCode, OITM.ItemName) ItemName, " & _
		"OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'UserText', OITM.ItemCode, OITM.UserText) UserText, " & _
		"PicturName " & _
		"from oitm where itemcode = N'" & saveHTMLDecode(Request("Item"), False) & "'"
set rs = conn.execute(sql)
		  If rs("PicturName") <> "" Then
		  	Pic = rs("PicturName")
		  Else
		  	Pic = "n_a.gif"
		  End If 
		  set rw = server.CreateObject("ADODB.recordset")
		  sql = "select top 15 tlog.LogNum, Convert(nvarchar(8),DocDate, 3) DocDate, " & _
		  		"OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OSLP', 'SlpName', OSLP.SlpCode, OSLP.SlpName) SlpName, " & _
  				"OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'SalUnitMsr', OITM.ItemCode, OITM.SalUnitMsr) SalUnitMsr, " & _
  				"Case When OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRD', 'CardName', TDOC.CardCode, TDOC.CardName) collate database_default is not null and LTRim(OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRD', 'CardName', TDOC.CardCode, TDOC.CardName)) collate database_default <> '' Then OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRD', 'CardName', TDOC.CardCode, TDOC.CardName) collate database_default Else TDOC.CardCode End CardName, " & _
  				"Case SaleType When 1 Then 'Un.' When 2 Then OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'SalUnitMsr', OITM.ItemCode, OITM.SalUnitMsr) collate database_default When 3 Then OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'SalPackMsr', OITM.ItemCode, OITM.SalPackMsr) collate database_default End As SaleType, " & _
  				"tdoc.CardCode, " & _
		  		"Quantity, Price, Convert(int,DocDate), SalPackUn, SaleType SaleType2, NumInSale " & _
		  		"from r3_obscommon..doc1 doc1 " & _
		  		"inner join r3_obscommon..tdoc tdoc on tdoc.lognum = doc1.lognum " & _
		  		"inner join r3_obscommon..tlog tlog on tlog.lognum = tdoc.lognum " & _
		  		"inner join OLKSalesLines T0 on T0.LogNum = doc1.Lognum and T0.LineNum = doc1.Linenum " & _
		  		"inner join oslp on oslp.slpcode = tdoc.slpcode " & _
		  		"inner join oitm on oitm.itemcode = doc1.itemcode collate database_default " & _
		  		"where doc1.ItemCode = N'" & saveHTMLDecode(Request("Item"), False) & "' and Status in ('R','H') order by 9 desc"

		  	set rw = conn.execute(sql)
%><table border="0" cellpadding="0"  bordercolor="#111111" width="100%" id="table7">
  <tr>
    <td width="335">
    <table border="0" cellpadding="0"  bordercolor="#111111" width="100%" id="table10">
      <tr>
        <td width="100%">
		<p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><%=getolkItemReport1LngStr("LtxtItmsComInOLK")%></font></b></td>
      </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td width="335">
    <table border="0" cellpadding="0"  bordercolor="#111111" width="100%" id="table8">
      <tr>
        <td width="10%" bgcolor="#D6EAFE"><b><font size="1" face="Verdana">&nbsp;<%=getolkItemReport1LngStr("LtxtReference")%>:</font></b></td>
        <td bgcolor="#D6EAFE"><b><font size="1" face="Verdana">
        <%=Request("Item")%>: <%=RS("ItemName")%></font></b></td>
      </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td>
    <table border="0" cellpadding="0" cellspacing="1"  bordercolor="#111111" width="100%" id="table9">
      <tr>
        <td bgcolor="#D2E9FF">
        <font face="Verdana" size="1">
    <b>#</b></font></td>
        <td bgcolor="#D2E9FF">
        <b><font face="Verdana" size="1"><%=getolkItemReport1LngStr("DtxtDate")%></font></b></td>
        <td bgcolor="#D2E9FF" colspan="2">
        <b><font face="Verdana" size="1"><%=getolkItemReport1LngStr("DtxtClient")%></font></b></td>
      </tr>
      <% do while not rw.eof %>
      <tr>
        <td bgcolor="#E6E6E6" style="color: #808080">
		<font face="Verdana" size="1" color=  "#777777"><b><%=RW("LogNum")%>&nbsp;</b></font></td>
        <td bgcolor="#E6E6E6"><font face="Verdana" size="1"><%=RW("DocDate")%>&nbsp;</font></td>
        <td bgcolor="#E6E6E6"<% If myApp.EnableUnitSelection Then %> colspan="2"<% End If %>>
		<font face="Verdana" size="1"><p><%=RW("CardName")%></font></td>
      </tr>
      <tr>
        <td bgcolor="#E2F0FE" ><font face="Verdana" size="1">
		<p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><% If rw("SaleType2") = 3 Then %><%=CDbl(RW("Quantity"))/CDbl(RW("SalPackUn"))%><% Else %><%=RW("Quantity")%><% End If %>&nbsp;</font></td>
        <td bgcolor="#E2F0FE" ><font face="Verdana" size="1">&nbsp;<%=RW("SLPName")%></font></td>
        <% If myApp.EnableUnitSelection Then %><td bgcolor="#E2F0FE" ><font face="Verdana" size="1"><%=RW("SaleType")%><% If Not myApp.UnEmbPriceSet And rw("SaleType2") = 3 Then %>&nbsp;<%=rw("SalUnitMsr")%><% If myApp.GetShowQtyInUn Then %>(<%=rw("NumInSale")%>)<% End If %><% End If %>&nbsp;</font></td><% End If %>
        <td bgcolor="#E2F0FE" ><font face="Verdana" size="1"><p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><%=myApp.MainCur%>&nbsp;<% If myApp.UnEmbPriceSet And rw("SaleType2") = 3 Then %><%=FormatNumber(CDbl(RW("Price"))*CDbl(rw("SalPackUn")),myApp.PriceDec)%><% Else %><%=FormatNumber(RW("Price"),myApp.PriceDec)%><% End If %></font></td>
      </tr>
      <% rw.movenext
      loop %>
      <tr>
        <td bgcolor="#E2F0FE"><font size="1" face="Verdana">&nbsp;<span lang="en-us">&nbsp;
		</span></font></td>
        <td bgcolor="#E2F0FE"><font size="1" face="Verdana">&nbsp;<span lang="en-us">&nbsp;
		</span></font></td>
        <% If myApp.EnableUnitSelection Then %><td bgcolor="#E2F0FE"><font size="1" face="Verdana">&nbsp;<span lang="en-us">&nbsp;
		</span></font></td><% End If %>
        <td bgcolor="#E2F0FE"><span lang="en-us"><font size="1" face="Verdana">&nbsp;
		</font></span></td>
      </tr>
    </table>
    </td>
  </tr>
  </table>
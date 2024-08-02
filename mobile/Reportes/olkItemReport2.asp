<head>
<style type="text/css">
.style1 {
	border-style: solid;
	border-width: 0;
}
.style2 {
	font-weight: bold;
	border-style: solid;
	border-width: 0;
}
.style3 {
	text-align: center;
}
</style>
</head>

<% addLngPathStr = "Reportes/" %>
<!--#include file="lang/olkItemReport2.asp" -->
<%
set rs = Server.CreateObject("ADODB.recordset")
sql = 	"select OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'ItemName', ItemCode, ItemName) ItemName, PicturName " & _
		"from oitm where itemcode = N'" & saveHTMLDecode(Request("Item"), False) & "'"
set rs = conn.execute(sql)

If rs("PicturName") <> "" Then Pic = rs("PicturName") Else Pic = "n_a.gif"

set rw = server.CreateObject("ADODB.recordset")
sql = 	"select top 10 DocNum, LineNum+1 LineNum, Convert(int,oinv.DocDate), Convert(nvarchar(10),oinv.docdate,3) DocDate, " & _
		"oinv.CardCode, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRD', 'CardName', OCRD.CardCode, OCRD.CardName) CardName, " & _
		"Case UseBaseUn When 'N' Then OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'SalUnitMsr', OITM.ItemCode, SalUnitMsr) collate database_default When 'Y' Then N'" & getolkItemReport2LngStr("DtxtUnit") & "' End MType, " & _
		"Quantity, Price " & _
		"from inv1 " & _
		"inner join oinv on oinv.docentry = inv1.docentry " & _
		"inner join oitm on oitm.itemcode = inv1.itemcode " & _
		"inner join OCRD on OCRD.CardCode = OINV.CardCode " & _
		"where inv1.ItemCode = N'" & saveHTMLDecode(Request("Item"), False) & "' and (targettype is null or targettype <> 14) "
	
If Not myAut.HasAuthorization(97) Then sql = sql & " and oinv.SlpCode = " & Session("vendid") & " "

sql = sql & "order by 3 desc"
set rw = conn.execute(sql)
%><table border="0" cellpadding="0"  bordercolor="#111111" width="100%" id="table2">
  <tr>
    <td width="100%">
    <table border="0" cellpadding="0"  bordercolor="#111111" width="100%" id="table6">
      <tr>
        <td width="97%">
		<p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><%=getolkItemReport2LngStr("LtxtLast10SalesOfItm")%></font></b></td>
      </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td width="100%">
    <table border="0" cellpadding="0"  bordercolor="#111111" width="100%" id="table3">
      <tr>
        <td width="10%" bgcolor="#D6EAFE"><b><font size="1" face="Verdana">&nbsp;<%=getolkItemReport2LngStr("LtxtReference")%>:</font></b></td>
        <td bgcolor="#D6EAFE"><b><font size="1" face="Verdana">
        <%=Request("Item")%>: <%=RS("ItemName")%></font></b></td>
      </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td>
    <table border="0" cellpadding="0" cellspacing="1"  bordercolor="#111111" width="100%" id="table5">
      <tr>
        <td bgcolor="#D2E9FF" style="width: 55px" class="style2">
        <font face="Verdana" size="1">
    	#</font></td>
        <td bgcolor="#D2E9FF" width="60">
        <b><font face="Verdana" size="1"><%=getolkItemReport2LngStr("DtxtDate")%></font></b></td>
        <td bgcolor="#D2E9FF">
        <b><font face="Verdana" size="1"><%=getolkItemReport2LngStr("DtxtDesc")%></font></b></td>
      </tr>
      <% do while not rw.eof %>
      <tr><form method="post" action="../cxcdocdetail.asp" target="_blank">
        <td bgcolor="#E6E6E6" style="width: 55px" class="style1">
        <font color="#FFFFFF">
        <input type="hidden" name="high" value="<%=RW("LineNum")%>">
        <input type="hidden" name="doctype" value="13">
        <input type="hidden" name="docnum" value="<%=RW("DocNum")%>">
		</font>
		<font face="Verdana" size="1"><%=RW("DocNum")%>(<%=RW("LineNum")%>)</font></td>
        </form>
        <td bgcolor="#E6E6E6" width="60" class="style3">
        <font face="Verdana" size="1">
        <%=RW("DocDate")%>&nbsp;</font></td>
        <td bgcolor="#E6E6E6">
        <font face="Verdana" size="1">
        <%=RW("CardName")%>&nbsp;</font></td>
      </tr>
      <tr>
        <td bgcolor="#E2F0FE" style="width: 55px" class="style1">
        <font size="1" face="Verdana">&nbsp;</font></td>
        <td bgcolor="#E2F0FE" width="60">
        <font face="Verdana" size="1">
        <p>
        <%=RW("Quantity")%>&nbsp;<%=RW("MType")%></font></td>
        <td bgcolor="#E2F0FE">
        <font face="Verdana" size="1">
		<p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><%=myApp.MainCur%>&nbsp;<%=FormatNumber(RW("Price"),myApp.PriceDec)%>&nbsp;</font></td>
      </tr>
      <% rw.movenext
      loop %>
      <tr>
        <td bgcolor="#E2F0FE" style="width: 55px" class="style1"><font size="1">&nbsp;</font></td>
        <td bgcolor="#E2F0FE" width="60"><font size="1">&nbsp;</font></td>
        <td bgcolor="#E2F0FE"><font size="1">&nbsp;</font></td>
      </tr>
    </table>
    </td>
  </tr>
  </table>
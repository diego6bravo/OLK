<% addLngPathStr = "cart/" %>
<!--#include file="lang/cart.asp" -->
<%
varx = 0
set rc = server.createobject("ADODB.RecordSet")
set rd = server.createobject("ADODB.RecordSet")
sql = "SELECT (select top 1 DirectRate from OADM) DirectRate, IsNull((select MaxDiscount from OLKAgentsAccess where SlpCode = " & Session("vendid") & "), 0) MaxLineDiscount, " & _
"(select SDKId from r3_obscommon..tcif where companydb = N'" & Session("olkdb") & "') SDKID " & _
"from olkcommon"
set rs = conn.execute(sql)

DirectRate = rs("DirectRate")
myLinesCount = getLinesCount()
SDKID = rs("SDKID")

If Session("useraccess") = "U" Then
	MaxDiscount = rs("MaxLineDiscount")
Else
	MaxDiscount = 100
End If

If myApp.ExpItems Then
	sqlAdd = ", IsNull((select Sum(LineTotal) from R3_ObsCommon..DOC3 T0 where LogNum = " & Session("RetVal") & "),0) Gastos, " & _
			"IsNull((select Sum(ExpenseTax) from ( " & _
			"select Case  " & _
			"	When 'PA' = 'PA' Then OLKCommon.dbo.DBOLKGetVatGroupRate" & Session("ID") & "(T1.VatGroupI, Convert(datetime,tdoc.DocDate,120)) " & _
			"	When 'PA' in ('MX', 'CR', 'GT', 'CL', 'US', 'CA') Then T1.RevFixSum " & _
			"	When 'PA' = 'IL' Then Case T1.TaxLiable When 'N' Then 0 Else (select VatPrcnt from OADM) End  " & _
			"End/100*LineTotal ExpenseTax " & _
			"from R3_ObsCommon..doc3 T0 " & _
			"inner join OEXD T1 on T1.ExpnsCode = T0.ExpnsCode " & _
			"where LogNum = " & Session("RetVal") & ") X0), 0) ExpenseTax "
End If

sql = "select cntctcode, comments, TDOC.CardCode, tdoc.CardName, E_Mail, object, tdoc.DocCur Currency, NumAtCard, tdoc.slpcode, IsNull(tdoc.DiscPrcnt, 0) DiscPrcnt, " & _
	  "(select top 1 VatPrcnt from oadm order by CurrPeriod desc) VatPrcnt, TDOC.ReserveInvoice, Confirm, Case When TDOC.DocDueDate is null Then 'N' Else 'Y' End VerfyDueDate, tdoc.DocDate " & sqlAdd & _
	  "from R3_ObsCommon..tdoc tdoc " & _
	  "inner join ocrd on ocrd.cardcode = tdoc.cardcode collate database_default " & _
	  "inner join r3_obscommon..tlog tlog on tlog.lognum = tdoc.lognum " & _
	  "inner join olkdocconf G0 on G0.objectcode = tlog.object " & _
	  "where tlog.lognum = " & Session("RetVal")
set rs = conn.execute(sql)
DocDate = rs("DocDate")
DiscPrcnt = CDbl(rs("DiscPrcnt"))

sql = 	"declare @UnEmbPriceSet char(1) set @UnEmbPriceSet = (select UnEmbPriceSet from olkcommon) " & _
		"select T0.LogNum, T0.ItemCode, T0.LineNum, " & _
		"OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'ItemName', T0.ItemCode, T1.ItemName) ItemName, " & _
		"OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'SalUnitMsr', T0.ItemCode, T1.SalUnitMsr) SalUnitMsr, " & _
		"OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'SalPackMsr', T0.ItemCode, T1.SalPackMsr) SalPackMsr, " & _
		"T1.NumInSale, T1.SalPackUn, T0.UnitPrice, " & _
		"T0.Price, T2.ManPrc, IsNull(IsNull(OLKCommon.dbo.DBOLKMyItmDiscPrice" & Session("ID") & "(N'" & saveHTMLDecode(Session("UserName"), False) & "', T1.ItemCode, getdate()),P0.Price), 0) SetPrice, " & _
		"Case SaleType When 1 Then T0.Quantity When 2 Then T0.Quantity When 3 Then T0.Quantity/SalPackUn End AS Quantity, T0.Currency, SaleType, " & _
		"Cast(T0.Price * T0.Quantity  As Decimal(20,2)) AS LineTotal, "
		
	If myApp.SVer < 8 Then
		sql = sql & "Case When " & SDKID & "LineMemo is not null Then 'Y' Else 'N' End HasMemo, "
	Else
		sql = sql & "Case When T10.LineText is not null Then 'Y' Else 'N' End HasMemo, "
	End If

	Select Case myApp.LawsSet 
		Case "PA", "AT", "AU", "BE", "CH", "CZ", "DE", "DK", "ES", "FI", "FR", "CN", "CY", "HU", "IT", "NL", "NO", "PL", "PT", "RU", "SE", "SK", "GB", "ZA"
			sql = sql & "Case VatGourpSa When 'V0' Then 0 Else Case T6.VatStatus When 'N' Then 0  " & _
							"Else OLKCommon.dbo.DBOLKGetVatGroupRate" & Session("ID") & "(IsNull(IsNull(T0.VatGroup, T6.ECVatGroup), T1.VatGourpSa), Convert(datetime,'" & SaveSqlDate(FormatDate(DocDate, False)) & "',120)) End End ITBM "
		Case "MX", "CL", "CR", "GT", "US", "CA", "BR"
			sql = sql & "Case VatLiable When 'N' Then 0 Else IsNull((select Rate from ostc where Code = T0.TaxCode collate database_default),0) End ITBM "
		Case "IL"
			sql = sql & "Case VatLiable When 'N' Then 0 When 'Y' Then (select top 1 VatPrcnt from oadm order by CurrPeriod desc) End As ITBM "
	End Select

sql = sql & " FROM R3_ObsCommon..DOC1 T0 " & _
		"INNER JOIN OITM T1 ON T1.ItemCode = T0.ItemCode collate database_default " & _
		"INNER JOIN OlkSalesLines T2 on T2.LogNum = T0.Lognum and T2.LineNum = T0.LineNum " & _
		"inner join OCRD T6 on T6.CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "' " & _
		"inner join ITM1 P0 on P0.ItemCode = T1.ItemCode and P0.PriceList = " & Session("plist") & " "
		
If myApp.SVer >= 8 Then
	sql = sql & "left outer join R3_ObsCommon..DOC10 T10 on T10.LogNum = T0.LogNum and T10.LineType = 'T' and T10.AfterLine = T0.LineNum "
End If
		
sql = sql & "Where T0.LogNum = " & Session("RetVal") & " "

		If myApp.EnableCartSum and Request("ViewMode") = "" Then
			If myApp.CartSumQty < myLinesCount and Request("document") <> "B" Then
				sql = sql & " and T0.LineNum >= (select Min(LineNum) from (select top " & myApp.CartSumQty & " LineNum from R3_ObsCommon..DOC1 X0 where LogNum = " & Session("RetVal") & " order by LineNum desc) T0) "
			Else
			If Request("String") <> "" Then
				arrSearchStr = Split(Request("String"), " ")
				sqlSearchFilter = ""
				For i = 0 to UBound(arrSearchStr)
					If sqlSearchFilter <> "" Then sqlSearchFilter = sqlSearchFilter & " or "
					sqlSearchFilter = sqlSearchFilter & " (T1.ItemCode like N'%" & arrSearchStr(i) & "%' or "
					sqlSearchFilter = sqlSearchFilter & " OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'ItemName', T0.ItemCode, T1.ItemName) like N'%" & arrSearchStr(i) & "%' or " & _
														"OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'FrgnName', T0.ItemCode, T1.frgnName) like N'%" & arrSearchStr(i) & "%') "				
				Next
				sql = sql  & " and ((" & sqlSearchFilter & ") or T1.CodeBars = N'" & Request("String") & "') "
			End If
			End If
		End If
			
		sql = sql & "Order By T0.LineNum desc"
rd.open sql, conn, 3, 1

If Not rd.Eof Then
	LineArr = ""
	do while not rd.eof
		If LineArr <> "" Then LineArr = LineArr & ", "
		LineArr = LineArr & rd("LineNum")
	rd.movenext
	loop
	rd.movefirst
End If

MaxDocDiscount = myApp.MaxDiscount
If Session("UserAccess") = "P" and not myApp.ApplyMaxDiscToSU Then MaxDocDiscount = 100
%>
<script language="javascript" src="cart/cart.js"></script>
<SCRIPT LANGUAGE="JavaScript">
var txtMaxDiscount = '<%=getcartLngStr("LtxtMaxDiscount")%>';
var MaxDiscount = '<%=FormatNumber(MaxDiscount, myApp.PercentDec)%>';
var MaxDocDisc = '<%=FormatNumber(MaxDocDiscount, myApp.PercentDec)%>';
var PercentDec = <%=myApp.PercentDec%>;
var PriceDec = <%=myApp.PriceDec%>;
var SumDec = <%=myApp.SumDec%>;
var QtyDec = <%=myApp.QtyDec%>;
var DocCur = '<%=rs("Currency")%>';
var formatDec = '<%=GetFormatDec()%>';
var txtValNumVal = '<%=getcartLngStr("DtxtValNumVal")%>';
var txtConfDelItm = '<%=getcartLngStr("LtxtConfDelItm")%>';
var txtValNumMaxVal = '<%=getcartLngStr("DtxtValNumMaxVal")%>';
var UnEmbPriceSet = <%=JBool(myApp.UnEmbPriceSet)%>;
var txtValDelDate = '<%=getcartLngStr("LtxtValDelDate")%>';
var txtValEnterValue = '<%=getcartLngStr("DtxtValEnterValue")%>';
var AgentSaleUnit = <%=myApp.GetSaleUnit%>;
function setTotal() 
{
var varSubTotal = 0;
var varTax = 0;
var discPrcnt = parseFloat(document.frmCart.DiscPrcnt.value);

<% if not rd.eof then
do while not rd.eof %>
varSubTotal = varSubTotal + parseFloat(myTrim(document.frmCart.LineTotal<%=rd("LineNum")%>.value.replace('<%=rd("Currency")%>','')));
varTax = varTax + parseFloat(myTrim(document.frmCart.LineTotal<%=rd("LineNum")%>.value.replace('<%=rd("Currency")%>','')))*(parseFloat(document.frmCart.ITMTax<%=RD("LineNum")%>.value)/100);
<% rd.movenext
loop 
rd.movefirst
end if
%>

<% If myApp.EnableCartSum and myApp.CartSumQty < myLinesCount and Request("ViewMode") = "" Then %>
varSubTotal = varSubTotal + parseFloat(document.frmCart.addTotal.value);
varTax = varTax + parseFloat(document.frmCart.addTax.value);
<% End If %>
varTax = varTax - (varTax*(discPrcnt/100));
<% If myApp.ExpItems Then %>
varTax = varTax + parseFloat(document.frmCart.ExpenseTax.value);
<% End If %>

varGastos = 0<% If myApp.ExpItems Then %> + parseFloat(myTrim(document.frmCart.Gastos.value.replace('<%=RS("Currency")%>','')))<% End If %>;

var discAmt = varSubTotal*(discPrcnt/100);
document.frmCart.DiscPrcntAmt.value = "<%=RS("Currency")%> " + FormatNumber(discAmt,<%=myApp.SumDec%>);
document.frmCart.SubTotal.value = "<%=RS("Currency")%> " + FormatNumber(varSubTotal,<%=myApp.SumDec%>);
document.frmCart.ITBM.value = "<%=RS("Currency")%> " + FormatNumber(varTax,<%=myApp.SumDec%>);
document.frmCart.importe.value = "<%=RS("Currency")%> " + FormatNumber(parseFloat(FormatNumber(varSubTotal,<%=myApp.SumDec%>))+parseFloat(FormatNumber(varTax,<%=myApp.SumDec%>))+varGastos-parseFloat(FormatNumber(discAmt,<%=myApp.SumDec%>)),<%=myApp.SumDec%>);

}
<% If Request("fastAddErr") = "Y" Then %>
<% Select Case Request("fastAddErrType")
Case "B" %>
alert('<%=Replace(getcartLngStr("LtxtFastAddErrBlock"), "{0}", Request("fastAddErrItm"))%>');
<% Case Else %>
alert('<%=Replace(getcartLngStr("LtxtFastAddErr"), "{0}", Request("fastAddErrItm"))%>');
<% End Select %>
<% ElseIf Request("err") = "disp" Then %>
alert('<%=getcartLngStr("LtxtErrItmInv")%>');
<% End If %>

<% 
set rDiscount = Server.CreateObject("ADODB.RecordSet")
sql = getDisItmPrcQry()
set rDiscount = conn.execute(sql)
If rDiscount.Eof Then %>
var itemVolDisc = null;
<% Else
ItemCode = ""
do while not rDiscount.eof
	If ItemCode <> rDiscount("ItemCode") Then 
		If itemVolDisc <> "" Then itemVolDisc = itemVolDisc & "{I}"
		itemVolDisc = itemVolDisc & rDiscount("ItemCode") & "{D}"
		ItemCode = rDiscount("ItemCode")
	Else
		If itemVolDisc <> "" Then itemVolDisc = itemVolDisc & "{S}"
	End If
	itemVolDisc = itemVolDisc & rDiscount("Amount") & "|" & CDbl(rDiscount("Price"))
rDiscount.movenext
loop %>
var itemVolDisc = '<%=itemVolDisc%>';
<% End If
set rDiscount = Nothing %>

</SCRIPT>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
        <tr>
          <td width="100%" bgcolor="#9BC4FF">
          <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><%=getcartLngStr("LtxtShopCart")%> - <% Select Case rs("Object")
          Case 17
          	Response.Write txtOrdr
          Case 23
          	Response.Write txtQuote
          Case 13
          	Select Case rs("ReserveInvoice")
          		Case "Y"
		          	Response.Write txtInvRes
          		Case Else
		          	Response.Write txtInv
		    End Select
          End Select%> #<%=Session("RetVal")%></font></b></td>
        </tr>
        <tr>
        	<td bgcolor="#9BC4FF" align="center">
				<form name="frmAddFast" action="cart/cartFastAddSubmit.asp" method="post" onsubmit="return valFastAdd();">
        		<input type="button" name="btnFastAddReset" value="X" style="font-family: Verdana; font-size: 10px; border: 1px solid #004ABB; background-color: #3E8BFF; height: 16px;" onclick="javascript:resetFastAdd();"><input type="text" name="txtFastAdd" value="" size="24" maxlength="24" style="font-family: Verdana; font-size: 10px;" ><br>
        		<select name="SaleUnit" size="1">
        		<% If Request("fastUnit") = "" Then 
        			If myApp.FastAddUnRem and Session("CurSaleType") <> "" Then
        				fastUnit = Session("CurSaleType")
        			Else
	        			fastUnit = myApp.GetSaleUnit 
	        		End If
        		Else 
        			fastUnit = CInt(Request("fastUnit"))
        		End If %>
        		<option value="1"><%=getcartLngStr("DtxtUnit")%></option>
        		<option <% If fastUnit = 2 Then %>selected<% End If %> value="2"><%=getcartLngStr("DtxtSalUnit")%></option>
        		<option <% If fastUnit = 3 Then %>selected<% End If %> value="3"><%=getcartLngStr("DtxtPackUnit")%></option></select><input type="number" min="0" step="<%=GetNumberStep(myApp.QtyDec)%>" name="txtFastAddQty" onclick="this.selectionStart=0;this.selectionEnd=this.value.length;" value="<% If Request("txtFastAddQty") = "" Then %><%=FormatNumber(1, myApp.QtyDec)%><% Else %><%=Request("txtFastAddQty")%><% End If %>" size="6" maxlength="20" style="font-family: Verdana; font-size: 10px; text-align: right;"><input type="submit" name="btnFastAdd" value="+" style="font-family: Verdana; font-size: 10px; border: 1px solid #004ABB; background-color: #3E8BFF; height: 16px;">
        		</form>
        	</td>
        </tr>
    <tr>
      <td bgcolor="#9BC4FF">
      <form method="POST" action="cart/cartupdate.asp" name="frmCart">
      <input type="hidden" name="VerfyDueDate" value="<%=rs("VerfyDueDate")%>">
      <input type="hidden" name="ObjCode" value="<%=rs("Object")%>">
        <table border="0" cellpadding="0" bordercolor="#111111" width="100%" id="AutoNumber1">
        <tr>
          <td width="100%" style="border-bottom-style: solid; border-bottom-width: 1px">
          <table border="0" cellpadding="0" bordercolor="#111111" width="100%">
            <tr>
              <td width="33%"><b><font size="1" face="Verdana"><%=getcartLngStr("DtxtFor")%>:</font></b></td>
              <td width="7%"><font size="2">
              <a href="operaciones.asp?cmd=datos&card=<%=CleanItem(rs("CardCode"))%>">
			<img border="0" src="images/<%=Session("rtl")%>flechaselec.gif"></a></font></td>
              <td width="50%"><input type="text" name="CardName" size="16" style="font-family: Verdana; font-size: 10px" value="<%=Replace(myHTMLEncode(rs("CardName")), """", "&quot;")%>" onclick="this.selectionStart=0;this.selectionEnd=this.value.length;"></td>
              <td width="10%">
               <% if rs("E_Mail") <> "" then %><a href="mailto:<%=rs("E_Mail")%>"><img border="0" src="cart/mail_icon.gif"></a><%end if%></td>
            </tr>
            <tr>
              <td width="33%"><b><font size="1" face="Verdana"><%=getcartLngStr("DtxtContact")%>:</font></b></td>
              <td width="57%" colspan="2">
        <select size="1" name="CntctCode" style="font-size:10px; width:100; font-family:Verdana">
        <%  
        set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKGetBPContacts" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@CardCode") = rs("CardCode")
        set rc = cmd.execute()
        do while not rc.eof %>
               <option <%If Not IsNull(rs("cntctcode")) then If CInt(rs("cntctcode")) = CInt(rc("cntctcode")) then response.write "selected"%> value="<%=rc("cntctcode")%>"><%=rc("Name")%></option>
               <%rc.movenext
               loop %>
               </select></td>
              <td width="10%">
              </td>
            </tr>
            <tr>
              <td width="100%" colspan="4">
          <table border="0" cellpadding="0" cellspacing="1"  bordercolor="#111111" width="100%" id="AutoNumber4" height="60">
            <tr>
              <td bgcolor="#95BFFF">
              <p align="center">&nbsp;</td>
              <td bgcolor="#75ACFF">
              <font size="1" face="Verdana">&nbsp;<b><%=getcartLngStr("LtxtSubTotal")%></b></font><b><font size="1" face="Verdana">:</font></b></td>
              <td bgcolor="#75ACFF" align="right"><b>
              <font size="1" face="Verdana" color="#FF0000">
                  <input disabled name="SubTotal" size="25" style="font-family: Verdana; font-size: 10px; color: #000000; text-align:right; " dir="ltr"></font></b></td>
       		</tr>
            <tr>
              <td bgcolor="#95BFFF">
              <p align="center">&nbsp;</td>
              <td bgcolor="#75ACFF"><nobr>
              <font size="1" face="Verdana">&nbsp;<b><%=getcartLngStr("DtxtDiscount")%><input name="DiscPrcnt" <% If Not myAut.HasAuthorization(91) Then %>readonly<% End If %> size="6" value="<%=FormatNumber(DiscPrcnt, myApp.PercentDec)%>" onclick="this.selectionStart=0;this.selectionEnd=this.value.length;" onchange="setDocDisc();"  style="font-family: Verdana; font-size: 10px; color: #000000; text-align:right; " dir="ltr" type="number" step="<%=GetNumberStep(myApp.PercentDec)%>"><input type="hidden" name="oldDiscPrcnt" value="<%=FormatNumber(DiscPrcnt, myApp.PercentDec)%>">%</b></font><b><font size="1" face="Verdana">:</font></b></nobr></td>
              <td bgcolor="#75ACFF" align="right"><b>
              <font size="1" face="Verdana" color="#FF0000">
                  <input name="DiscPrcntAmt" size="25" style="font-family: Verdana; font-size: 10px; color: #000000; text-align:right; " dir="ltr" disabled></font></b></td>
       		</tr>
            <% If myApp.ExpItems Then %>
			<tr>
              <td bgcolor="#95BFFF">
              <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><font size="2">
              <a href="operaciones.asp?cmd=cartExp"><img border="0" src="images/<%=Session("rtl")%>flechaselec.gif"></a>&nbsp; </font></td>
              <td bgcolor="#75ACFF">
              <font size="1" face="Verdana"><b>&nbsp;<%=getcartLngStr("LtxtExpenses")%>:</b></font></td>
              <td bgcolor="#75ACFF" align="right"><b>
              <font size="1" face="Verdana" color="#FF0000">
              	<input type="hidden" name="ExpenseTax" value="<%=rs("ExpenseTax")%>">
				<input disabled name="Gastos" size="25" style="font-family: Verdana; font-size: 10px; color: #000000; text-align:right; " value="<%=rs("Currency")%>&nbsp;<%=Replace(FormatNumber(rs("Gastos"),myApp.SumDec), GetFormatSep(), "")%>" dir="ltr"></font></b></td>
            	</tr>
            	<% End If %>
				<tr>
              <td bgcolor="#95BFFF">
              <p align="center">&nbsp;</td>
              <td bgcolor="#75ACFF">
              <font size="1" face="Verdana">&nbsp;<b><%=txtTax%></b></font><b><font size="1" face="Verdana">:</font></b></td>
              <td bgcolor="#75ACFF" align="right"><b>
              <font size="1" face="Verdana" color="#FF0000">
				<input disabled name="ITBM" size="25" style="font-family: Verdana; font-size: 10px; color: #000000; text-align:right; " dir="ltr"></font></b></td>
            	</tr>
            <tr>
              <td bgcolor="#95BFFF">
              <p align="center">&nbsp;</td>
              <td bgcolor="#75ACFF">
              <font size="1" face="Verdana">&nbsp;<b><%=getcartLngStr("DtxtTotal")%></b></font><b><font size="1" face="Verdana">:</font></b></td>
              <td height="9" bgcolor="#75ACFF" align="right"><b>
              <font size="1" face="Verdana" color="#FF0000">
              <input disabled type="text" name="importe" size="25" style="font-family: Verdana; font-size: 10px; color: #000000; text-align:right; float:right" dir="ltr"></font></b></td>
            </tr>
            </table>
              </td>
            </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td width="100%" style="font-size: 3px; border-left-width: 1px; border-right-width: 1px; border-top: 1px solid #FF9933; border-bottom-width: 1px">&nbsp;</td>
        </tr>
        <TR>
        <td>
          <div align="center">
            <center>
            <table border="0" cellpadding="0" cellspacing="1" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber5">
              <tr>
              	<td width="37">
                <p align="center">
                <% If myApp.EnableCartSum Then %>
                <a href="operaciones.asp?cmd=searchCart">
                <img border="0" src="images/cartlupa.gif"></a><% Else %>&nbsp;<% End If %></td>
				<td width="29">
                <a href="operaciones.asp?cmd=slistsearch">
                <img border="0" src="cart/search_icon.gif"></a></td>
              	<td>&nbsp;</td>
                <td width="25">
                <p align="center">
                <input border="0" src="images/save_icon.gif" name="I1" type="image"></td>
                <td width="31">&nbsp;</td>
				<% If not rd.eof Then %>
                <td width="37">
                <p align="center"><input type="image" name="I2" border="0" src="images/ok_icon.gif" onclick="javascript:return valFrm();"></td><% End If %>
				<td width="29">
                <p align="center"><a href="operaciones.asp?cmd=cartcancel&c1=<%=CleanItem(rs("CardCode"))%>"><img border="0" src="images/x_icon.gif"></a></td>
              </tr>
              <tr>
              	<td colspan="7" style="padding-top: 2px;">
              	<table cellpadding="0" cellspacing="0" border="0" width="100%">
              		<tr>
              			<td><input type="button" value="<%=txtBasketMinRep%>" name="btnCartCP" onclick="javascript:window.location.href='operaciones.asp?cmd=cart_cp'" style="font-family: Verdana; font-size: 10px; border: 1px solid #004ABB; background-color: #3E8BFF"></td>
              			<td align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><input type="button" value="<%=getcartLngStr("LtxtDetails")%>" name="btnCartDetails" onclick="javascript:window.location.href='operaciones.asp?cmd=cartopt'" style="font-family: Verdana; font-size: 10px; border: 1px solid #004ABB; background-color: #3E8BFF"></td>
              		</tr>
              	</table>
	            </td>
              </tr>
            </table>
            </center>
          </div>
        </td></TR>
        <tr>
          <td width="100%">
          <table border="0" cellpadding="0" cellspacing="1"  bordercolor="#111111" width="100%">
            <tr>
              <td bgcolor="#75ACFF">
              <b><font size="1" face="Verdana"><%=getcartLngStr("DtxtCode")%> | <%=getcartLngStr("DtxtDescription")%></font></b></td>
              <td bgcolor="#75ACFF" align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><% If Not rd.Eof Then %><input type="image" name="btnDel" border="0" src="cart/x_botoom.gif" onclick="return valDelItems();"><% End If %></td>
            </tr>
            <%
            If myApp.EnableCartSum and myApp.CartSumQty < myLinesCount and Request("ViewMode") = "" and not rd.eof Then
			sql = GetLineSumQty
			set rctn = conn.execute(sql) %>
			<tr>
				<td colspan="2" bgcolor="#95BFFF" align="center"><b>
				<font size="1" face="Verdana"><% If Request("document") <> "B" Then %><%=getcartLngStr("LtxtSumLines")%> - <%=myLinesCount-myApp.CartSumQty%><% Else %><%=getcartLngStr("LtxtFilterCartNote")%><% End If %></font></b></td>
			</tr>
			<tr>
				<td bgcolor="#95BFFF" colspan="2" align="right"><b><font size="1" face="Verdana"><%=getcartLngStr("LtxtTotalSummary")%>: <nobr><%=rs("Currency")%>&nbsp;<%=FormatNumber(rctn(0), myApp.SumDec)%></nobr></font></b>
		  		<input type="hidden" name="addTotal" value="<%=rctn(0)%>">
		  		<input type="hidden" name="addTax" value="<%=rctn(1)%>">
				</td>
			</tr>			
	  		<% End If %>
            <% do while not rd.eof
            discItem = IsDiscItem(rd("ItemCode")) %>
            <tr>
              <td width="100%" colspan="2" bgcolor="#95BFFF">
              <table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tr>
					<td valign="bottom" colspan="4">
              		<table border="0" width="100%" cellspacing="0" cellpadding="0">
						<tr>
							<td>
							<table cellpadding="0" cellspacing="0" border="0" width="100%">
								<tr>
									<td><b><font size="1" face="Verdana"><a href="operaciones.asp?cmd=itemdetails&Item=<%=Server.URLEncode(rd("ItemCode"))%>"><%=RD("ItemCode")%></a><br><%=Server.HTMLEncode(RD("ItemName"))%></font></b>
									</td>
									<% If discItem Then %><td width="23"><a href="javascript:window.location.href='operaciones.asp?cmd=itemVolRep&Item=<%=Server.URLEncode(rd("ItemCode"))%>&un=' + document.frmCart.selUn<%=RD("LineNum")%>.value;"><img border="0" src="images/foco_icon.gif" width="23" height="22"></a></td><% End If %>
								</tr>
							</table>
							</td>
							<td width="10" align="center" valign="middle">
							<input type="checkbox" value="<%=rd("LineNum")%>" id="chkDel" name="chkDel"></td>
						</tr>
					</table>
		            <input type="hidden" name="NumInSale<%=RD("LineNum")%>" value="<%=RD("NumInSale")%>">
		            <input type="hidden" name="SalPackUn<%=RD("LineNum")%>" value="<%=RD("SalPackUn")%>">
		            <input type="hidden" name="un<%=RD("LineNum")%>" value="<%=RD("SaleType")%>">
		            <input type="hidden" name="ITMTax<%=RD("LineNum")%>" value="<%=RD("ITBM")%>">
		            <input type="hidden" name="Currency<%=rd("LineNum")%>" value="<%=RD("Currency")%>">
				    <input type="hidden" name="ManPrc<%=rd("LineNum")%>" value="<%=rd("ManPrc")%>">
					</td>
				</tr>
				<tr>
					<td valign="bottom">
              		<b><font size="1" face="Verdana"><%=getcartLngStr("LtxtQty")%></font></b></td>
					<td>
					<table cellpadding="0" cellspacing="0" border="0">
						<tr>
							<% If myAut.HasAuthorization(92) Then %>
							<td valign="bottom">
							<a href='operaciones.asp?cmd=cartEditLine&amp;LineNum=<%=rd("LineNum")%>'><img src="images/expand<% If rd("HasMemo") = "Y" Then %>blue<% End If %>.gif" border="0"></a>
							</td><% End If %>
							<td>
							<input dir="ltr" style="FONT-SIZE: 10px; TEXT-ALIGN: right; " type="number" size="10" value="<%=FormatNumber(RD("Quantity"), myApp.QtyDec)%>" name="T<%=RD("LineNum")%>" onclick="this.selectionStart=0;this.selectionEnd=this.value.length;" onchange="javascript:chkQty(this, '<%=RD("Quantity")%>', <%=myApp.QtyDec%>, '<%=Replace(Replace(RD("ItemCode"), "'", "\'"), """", """""") %>', document.frmCart.ManPrc<%=rd("LineNum")%>, document.frmCart.LineTotal<%=RD("LineNum")%>, document.frmCart.price<%=RD("LineNum")%>, document.frmCart.NumInSale<%=RD("LineNum")%>.value, document.frmCart.SalPackUn<%=RD("LineNum")%>.value, document.frmCart.un<%=RD("LineNum")%>.value, document.frmCart.Currency<%=rd("LineNum")%>.value, document.frmCart.SetPrice<%=rd("LineNum")%>.value);" min="0" step="<%=GetNumberStep(myApp.QtyDec)%>">
					        </td>
						</tr>
					</table>
                  	</td>
					<td> <font size="1" face="Verdana">
                  	<b><%=getcartLngStr("DtxtPrice")%></b></font></td>
					<td>
					<input <% If Not myAut.HasAuthorization(68) Then %> readonly <% End If %> class="input" value="<% If Not myApp.UnEmbPriceSet and rd("SaleType") = 3 Then %><%=Replace(FormatNumber(CDbl(RD("Price"))*CDbl(rd("SalPackUn")),myApp.PriceDec),GetFormatSep(),"")%><% Else %><%=Replace(FormatNumber(CDbl(RD("Price")),myApp.PriceDec),GetFormatSep(),"")%><% End If %>" name="price<%=RD("LineNum")%>" size="16" dir="ltr" style="font-size: 10px; text-align:right; font-family:Verdana" onclick="this.selectionStart=0;this.selectionEnd=this.value.length;" onchange="chkPrice(this, '<%=FormatNumber(CDbl(RD("Price")),myApp.PriceDec)%>', <%=myApp.PriceDec%>, document.frmCart.ManPrc<%=rd("LineNum")%>, document.frmCart.LineTotal<%=RD("LineNum")%>, document.frmCart.T<%=RD("LineNum")%>, document.frmCart.SalPackUn<%=RD("LineNum")%>, document.frmCart.un<%=RD("LineNum")%>, document.frmCart.Currency<%=rd("LineNum")%>, document.frmCart.UnitPrice<%=rd("LineNum")%>);" type="number" min="0" step="<%=GetNumberStep(myApp.PriceDec)%>">
					<input type="hidden" name="SetPrice<%=rd("LineNum")%>" value="<%=Replace(FormatNumber(CDbl(rd("SetPrice")), myApp.PriceDec),GetFormatSep(),"")%>">
					<input type="hidden" name="UnitPrice<%=rd("LineNum")%>" value="<%=Replace(FormatNumber(CDbl(rd("UnitPrice")), myApp.PriceDec),GetFormatSep(),"")%>">
					</td>
				</tr>
				<tr>
					<td valign="bottom"><% If myApp.EnableUnitSelection Then %>
              		<b><font size="1" face="Verdana"><%=getcartLngStr("DtxtUnit")%></font></b><% Else %>&nbsp;<% End If %>
					</td>
					<td><% If myApp.EnableUnitSelection Then %>
                  	<select size="1" class="input" name="selUn<%=RD("LineNum")%>" style="font-size: 10px; width:75; font-family:Verdana" onchange="javascript:changeUnEmb(this, document.frmCart.price<%=RD("LineNum")%>,document.frmCart.T<%=RD("LineNum")%>,document.frmCart.un<%=RD("LineNum")%>, document.frmCart.LineTotal<%=RD("LineNum")%>, document.frmCart.SalPackUn<%=RD("LineNum")%>.value, document.frmCart.NumInSale<%=RD("LineNum")%>.value, document.frmCart.Currency<%=rd("LineNum")%>.value)">
					<option value="1" <% If RD("SaleType") = "1" Then Response.Write "selected"%>><%=getcartLngStr("DtxtUnit")%></option>
					<option value="2" <% If RD("SaleType") = "2" Then Response.Write "selected"%>><%=RD("SalUnitMsr")%><% If myApp.GetShowQtyInUn Then %>(<%=RD("NumInSale")%>)<% End If %></option>
					<option value="3" <% If RD("SaleType") = "3" Then Response.Write "selected"%>><%=RD("SalPackMsr")%><% If myApp.GetShowQtyInUn Then %>(<%=RD("SalPackUn")%>)<% End If %></option>
					</select><% Else %>&nbsp;<input type="hidden" name="selUn<%=RD("LineNum")%>" value="<%=RD("SaleType")%>"><% End If %></td>
					<td><b><font size="1" face="Verdana"> <%=getcartLngStr("DtxtTotal")%></font></b></td>
					<td>
					<input class="input" value="<%=RD("Currency")%>&nbsp;<%=Replace(FormatNumber(RD("LineTotal"),myApp.SumDec),GetFormatSep(),"")%>" name="LineTotal<%=RD("LineNum")%>" size="16" dir="ltr" style="font-size: 10px; text-align:right; font-family:Verdana" onclick="this.selectionStart=0;this.selectionEnd=this.value.length;" disabled></td>
				</tr>
				</table>
				</td>
            </tr>
		  <% If Request(RD("LineNum")) = "Y" Then %>
		  <script language="javascript">
		  alert('<%=getcartLngStr("LtxtItmInvErr")%>'.replace('{0}', '<%=Replace(Rd("ItemCode"), "'", "\'")%>'));
		  </script>
		  <% end if
            rd.movenext
            loop
            If myApp.EnableCartSum and myApp.CartSumQty < myLinesCount Then %>
			<tr>
				<td colspan="2">
            <%  If Request("ViewMode") = "" Then
              	viewCmd = "all"
              	viewBtnStr = "" & getcartLngStr("LtxtViewAll") & ""
              Else
              	If Request("ViewMode") = "all" Then
                  	viewCmd = ""
                  	viewBtnStr = "" & getcartLngStr("LtxtViewSumm") & ""
                Else
                  	viewCmd = "all"
                  	viewBtnStr = "" & getcartLngStr("LtxtViewAll") & ""
                End If
              End If %>
            <input type="hidden" name="oldViewMode" value="<%=Request("ViewMode")%>">
            <input type="submit" value="<%=viewBtnStr%>" name="btnViewAll" style="font-family: Verdana; font-size: 10px; border: 1px solid #004ABB; background-color: #3E8BFF">
            <input type="hidden" name="ViewMode" value="<%=Request("ViewMode")%>">
            <% If Request("document") = "B" Then %>&nbsp;<input type="button" name="btnClearFilter" value="<%=getcartLngStr("LtxtClearFilter")%>" onclick="window.location.href='?cmd=cart'" style="font-family: Verdana; font-size: 10px; border: 1px solid #004ABB; background-color: #3E8BFF"><% End If %></td>
			</tr>
            <% End If %>
            </table>
          </td>
        </tr>
      </table>
      <input type="hidden" name="document" value="<%=Request("document")%>">
      <input type="hidden" name="string" value="<%=Request("string")%>">
      </form>
      </td>
    </tr>
    </table>
  </center>
</div>
<script>setTotal()</script>
<% set rc = nothing
Function getLinesCount()
	sql = "SELECT Count('A') from R3_OBSCommon..DOC1 T0  " & _
	"inner join OITM T1 on T1.ItemCode = T0.ItemCode collate database_default  " & _
	"Where T0.LogNum = " & Session("RetVal")
	set rCount = Server.CreateObject("ADODB.RecordSet")
	set rCount = conn.execute(sql)
	getLinesCount = rCount(0)
End Function

Private Function GetLineSumQty()
	sqlStr = 				"declare @LogNum int set @LogNum = " & Session("RetVal") & " " & _
							"declare @MainCur nvarchar(3) set @MainCur = (select top 1 MainCurncy from oadm) " & _
							"select Sum(LineTotal) LineTotal, Sum(TaxTotal) TaxTotal " & _
							"from ( " & _
							"select ((T0.Price * T0.Quantity)" &  GetRateFunc(1) & "Case When T0.Currency = @MainCur Then 1 Else " & _
							"	(select Rate from ORTT where DateDiff(day,T2.DocDate,RateDate) = 0 and Currency = T0.Currency collate database_default) End)" &  GetRateFunc(2) & " " & _
							"	Case When T2.DocCur = @MainCur Then 1 Else  " & _
							"	(select Rate from ORTT where DateDiff(day,T2.DocDate,RateDate) = 0 and Currency = T2.DocCur collate database_default) End LineTotal, " & _
							"   (((T0.Price * T0.Quantity)" &  GetRateFunc(1) & "Case When T0.Currency = @MainCur Then 1 Else " & _
							"	(select Rate from ORTT where DateDiff(day,T2.DocDate,RateDate) = 0 and Currency = T0.Currency collate database_default) End)" &  GetRateFunc(2) & " " & _
							"	Case When T2.DocCur = @MainCur Then 1 Else  " & _
							"	(select Rate from ORTT where DateDiff(day,T2.DocDate,RateDate) = 0 and Currency = T2.DocCur collate database_default) End)* "
				
	Select Case myApp.LawsSet
		Case "MX", "CL", "CR", "GT", "US", "CA", "BR"
			sqlStr = sqlStr & 	"IsNull((select Rate/100 from OSTC where Code = T0.TaxCode collate database_default) ,0)"
		Case "PA", "AT", "AU", "BE", "CH", "CZ", "DE", "DK", "ES", "FI", "FR", "CN", "CY", "HU", "IT", "NL", "NO", "PL", "PT", "RU", "SE", "SK", "GB", "ZA"
		  	sqlStr = sqlStr & 	"((Case VatGourpSa When 'V0' Then 0 Else Case T3.VatStatus When 'N' Then 0 Else " & _
								"OLKCommon.dbo.DBOLKGetVatGroupRate" & Session("ID") & "(IsNull(IsNull(T0.VatGroup, T3.ECVatGroup), T1.VatGourpSa), Convert(datetime,'" & SaveSqlDate(FormatDate(DocDate, False)) & "',120)) End End)/100) "
		Case "IL" 
			sqlStr = sqlStr & "Case VatLiable When 'N' Then 0 When 'Y' Then (select top 1 VatPrcnt from oadm order by CurrPeriod desc) End "
	End Select
		
	sqlStr = sqlStr & 		" TaxTotal "
	
	sqlAddStr = ""
	
	sqlStr = sqlStr & 		"from r3_obscommon..doc1 T0  " & _
							"inner join oitm T1 on T1.itemcode = T0.itemcode collate database_default " & _
							"inner join R3_ObsCommon..TDOC T2 on T2.LogNum = T0.LogNum " & _
							"inner join ocrd T3 on T3.CardCode = T2.CardCode collate database_default " & _
							"left outer join R3_ObsCommon..DOC2 T4 on T4.LogNum = T0.LogNum and T4.LineNum = T0.LineNum "

	sqlStr = sqlStr & 		"where T0.LogNum = @LogNum and "
	
	If Request("document") <> "B" Then
		sqlStr = sqlStr & "T0.LineNum < (select Min(LineNum) from (select top " & myApp.CartSumQty & " LineNum from R3_ObsCommon..DOC1 X0 where LogNum = " & Session("RetVal") & " order by LineNum desc) T0) "
	Else
		If Request("String") <> "" Then
			arrSearchStr = Split(Request("String"), " ")
			sqlSearchFilter = ""
			For i = 0 to UBound(arrSearchStr)
				If sqlSearchFilter <> "" Then sqlSearchFilter = sqlSearchFilter & " or "
				sqlSearchFilter = sqlSearchFilter & " (ItemCode like N'%" & arrSearchStr(i) & "%' or ItemName like N'%" & arrSearchStr(i) & "%'" & _
	  											" or frgnName like N'%" & arrSearchStr(i) & "%') "
			Next
			sqlStr = sqlStr & " T0.ItemCode not in (select ItemCode collate database_default from OITM where ((" & sqlSearchFilter & ") or CodeBars = N'" & Request("String") & "')) "
		End If		
	End If

	sqlStr = sqlStr & 		"Group By T0.LineNum, T0.Price, T0.Quantity, T0.Currency, T2.DocDate, T0.Currency, T2.DocCur, T2.DocDate, T3.VatStatus, T0.VatGroup, T3.ECVatGroup, T1.VatGourpSa, T0.UseBaseUn, T1.NumInSale, T1.ManBtchNum "

	Select Case myApp.LawsSet 
		Case "MX", "CL", "CR", "GT", "US", "CA", "BR"
			sqlStr = sqlStr & 	", T0.TaxCode "
	End Select
	
	sqlStr = sqlStr & 		") T0 "
	GetLineSumQty = sqlStr
End Function

Function GetRateFunc(ByVal i)
	Select Case DirectRate
		Case "Y"
			Select Case i
				Case 1
					GetRateFunc = "*"
				Case 2
					GetRateFunc = "/"
			End Select
		Case "N"
			Select Case i
				Case 1
					GetRateFunc = "/"
				Case 2
					GetRateFunc = "*"
			End Select
	End Select
End Function

Function getDisItmPrcQry()
	sqlStr = "select T0.ItemCode, T0.Amount,  " & _
	"Case When AutoUpdt = 'N' Then T0.Price  " & _
	"			Else T2.Price-(T2.Price*T0.Discount/100) " & _
	"	End Price " & _
	"from spp2 T0 " & _
	"inner join spp1 T1 on T1.ItemCode = T0.ItemCode and T1.CardCode = T0.CardCode and T1.LineNum = T0.SPP1LNum " & _
	"inner join ITM1 T2 on T2.ItemCode = T0.ItemCode and T2.PriceList = T1.ListNum " & _
	"where (T0.CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "' or T0.CardCode = N'*" & Session("plist") & "' and not exists(select 'A' " & _
	"from spp2 Z0 " & _
	"inner join spp1 Z1 on Z1.ItemCode = Z0.ItemCode and Z1.CardCode = Z0.CardCode and Z1.LineNum = Z0.SPP1LNum " & _
	"inner join ITM1 Z2 on Z2.ItemCode = Z0.ItemCode and Z2.PriceList = Z1.ListNum " & _
	"where (Z0.CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "' and Z0.ItemCode = T0.ItemCode) " & _
	"and " & _
	"  (Z1.FromDate is null or DateDiff(day,getdate(),Z1.FromDate) <= 0) and " & _
	"  (Z1.ToDate is null or DateDiff(day,getdate(),Z1.ToDate) >= 0) " & _
	"and Z0.ItemCode in  " & _
	"  (	select ItemCode collate database_default " & _
	"	from R3_ObsCommon..DOC1 D0 " & _
	"	where LogNum = " & Session("RetVal") & " " & _
	"	and D0.LineNum >= " & _
	"	(select Min(LineNum) from " & _
	"		(select top 10 LineNum  " & _
	"		from R3_ObsCommon..DOC1 X0  " & _
	"		where LogNum = D0.LogNum  " & _
	"		order by LineNum desc) V0) " & _
	"  ))) " & _
	"and " & _
	"  (T1.FromDate is null or DateDiff(day,getdate(),T1.FromDate) <= 0) and " & _
	"  (T1.ToDate is null or DateDiff(day,getdate(),T1.ToDate) >= 0) " & _
	"and T0.ItemCode in  " & _
	"  (	select ItemCode collate database_default " & _
	"	from R3_ObsCommon..DOC1 D0 " & _
	"	where LogNum = " & Session("RetVal") & " " & _
	"	and D0.LineNum >= " & _
	"	(select Min(LineNum) from " & _
	"		(select top 10 LineNum  " & _
	"		from R3_ObsCommon..DOC1 X0  " & _
	"		where LogNum = D0.LogNum  " & _
	"		order by LineNum desc) V0) " & _
	"  ) "
	getDisItmPrcQry = sqlStr
End Function

Function IsDiscItem(Item)
	retVal = False
	If itemVolDisc <> "" Then
		arrItemVol = Split(itemVolDisc, "{I}")
		For i = 0 to UBound(arrItemVol)
			If Item = Split(arrItemVol(i), "{D}")(0) Then
				retVal = True
				Exit For
			End If
		Next
	End If
	IsDiscItem = retVal
End Function

 %>
 
 <script language="javascript">
 
function onScan(ev){
var scan = ev.data;
	document.frmAddFast.txtFastAdd.value = scan.value;
	document.frmAddFast.submit();
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
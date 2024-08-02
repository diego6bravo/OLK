<!--#include file="myHTMLEncode.asp" -->
<!--#include file="lang/myFunctions.asp" -->

<head>
<style type="text/css">
.functionStyle1 {
	font-size: xx-small;
	color: #31659C;
}
</style>
</head>

<% If Not HideFunctionTitle Then %><b><font face="Verdana" size="1" color="#31659C">
<%=getmyFunctionsLngStr("DtxtFunctions")%>:<br>
</font></b><% End If %>
<font size="1" color="#4783C5" face="Verdana">
<table cellpadding="0" border="0" width="100%">
	<tr>
		<td class="functionStyle1"><strong><%=getmyFunctionsLngStr("DtxtFunction")%></strong></td>
		<td class="functionStyle1"><strong><%=getmyFunctionsLngStr("DtxtDescription")%></strong></td>
		<td class="functionStyle1"><strong><%=getmyFunctionsLngStr("DtxtExample")%></strong></td>
	</tr>
	<tr class="<%=functionClass%>">
		<td>OLKCode</td>
		<td><%=getmyFunctionsLngStr("LtxtNumCod")%></td>
		<td><span dir="ltr">OLKCode('H', 1560.04, 2)</span></td>
	</tr>
	<tr class="<%=functionClass%>">
		<td>OLKDocTotal</td>
		<td><%=getmyFunctionsLngStr("LtxtDocTotal")%></td>
		<td><span dir="ltr">OLKDocTotal(@LogNum)</span></td>
	</tr>
	<tr class="<%=functionClass%>">
		<td>OLKDocTotalBox</td>
		<td><%=getmyFunctionsLngStr("LtxtTotalBox")%></td>
		<td><span dir="ltr">OLKDocTotalBox(@LogNum)</span></td>
	</tr>
	<tr class="<%=functionClass%>">
		<td>OLKDocTotalVol</td>
		<td><%=getmyFunctionsLngStr("LtxtTotalVol")%></td>
		<td><span dir="ltr">OLKDocTotalVol(@LogNum)</span></td>
	</tr>
	<tr class="<%=functionClass%>">
		<td>OLKDocTotalWeight</td>
		<td><%=getmyFunctionsLngStr("LtxtTotalWeight")%></td>
		<td><span dir="ltr">OLKDocTotalWeight(@LogNum)</span></td>
	</tr>
	<tr class="<%=functionClass%>">
		<td>OLKEncodeBreakLines</td>
		<td><%=getmyFunctionsLngStr("LtxtEncodeBreakLines")%></td>
		<td><span dir="ltr">OLKEncodeBreakLines(TextField)</span></td>
	</tr>
	<tr class="<%=functionClass%>">
		<td>OLKDateFormat</td>
		<td><%=getmyFunctionsLngStr("LtxtDateFormat")%></td>
		<td><span dir="ltr">OLKDateFormat('2012-12-31')</span></td>
	</tr>
	<tr class="<%=functionClass%>">
		<td>OLKFormatNumber</td>
		<td><%=getmyFunctionsLngStr("LtxtNumFormat")%></td>
		<td><span dir="ltr">OLKFormatNumber(1599.95, 2)</span></td>
	</tr>
	<tr class="<%=functionClass%>">
		<td>OLKGetCufv</td>
		<td><%=getmyFunctionsLngStr("LtxtGetCufv")%></td>
		<td><span dir="ltr">OLKGetCufv('TableID', 'AliasID', 'Value')</span></td>
	</tr>
	<tr class="<%=functionClass%>">
		<td>OLKGetTrans</td>
		<td><%=getmyFunctionsLngStr("LtxtGetTrans")%></td>
		<td><span dir="ltr">OLKGetTrans(@LanID, 'OTBL', 'COLID', T0.COLCode, T0.COLNam)</span></td>
	</tr>
	<tr class="<%=functionClass%>">
		<td>OLKInv</td>
		<td><%=getmyFunctionsLngStr("LtxtInv")%></td>
		<td><span dir="ltr">OLKInv(@ItemCode, @dbName, @WhsCode)</span></td>
	</tr>
	<tr class="<%=functionClass%>">
		<td>OLKItemInvVal</td>
		<td><%=getmyFunctionsLngStr("LtxtItemInvVal")%></td>
		<td><span dir="ltr">OLKItemInvVal(@ItemCode, @WhsCode, @dbName, @LogNum | -1, @LineNum | -1)</span></td>
	</tr>
	<tr class="<%=functionClass%>">
		<td>OLKNumIn</td>
		<td><%=getmyFunctionsLngStr("LtxtNumIn")%></td>
		<td><span dir="ltr">OLKNumIn(@OnHand, @NumIn)</span></td>
	</tr>
	<tr class="<%=functionClass%>">
		<td>OLKSplit</td>
		<td><%=getmyFunctionsLngStr("LtxtSplit")%></td>
		<td><span dir="ltr">OLKSplit('String', 'Delimiter')</span></td>
	</tr>
</table>
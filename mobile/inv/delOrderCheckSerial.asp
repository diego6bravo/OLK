<% addLngPathStr = "inv/" %>
<!--#include file="lang/delOrderCheckSerial.asp" -->

<head>
<style type="text/css">
.style1 {
	background-color: #FFFF00;
}
.style2 {
	background-color: #008000;
	color: #FFFF00;
}
.style5 {
	text-align: center;
}
</style>
</head>

<% 

sql = "declare @DocType char(1) set @DocType = (select DocType from OLKInOutSettings where ObjectCode = " & Session("ObjCode") & " and Type = '" & Session("Type") & "') " & _
"select T0.ItemCode, T0.CodeBars, T0.ItemName, (select ChkSerial from OLKInOutSettings where ObjectCode = " & Session("ObjCode") & " and Type = '" & Session("Type") & "') ChkSerial, " & _
"IsNull((select Sum(Case X0.UseBaseUn When 'N' Then X0.Quantity*Case @DocType When 'S' Then X1.NumInSale When 'P' Then X1.NumInBuy End When 'Y' Then X0.Quantity End) " & _
"from R3_ObsCommon..DOC1 X0 " & _
"inner join OITM X1 on X1.ItemCode = X0.ItemCode collate database_default " & _
"where X0.LogNum = " & Session("IORetVal") & " and X0.ItemCode = T0.ItemCode collate database_default and X0.WhsCode = N'" & Session("bodega") & "'), 0) ReqQty, " & _
"IsNull((select Count('') from R3_ObsCommon..DOC4 where LogNum = " & Session("IORetVal") & " and LineNum in  " & _
"(select LineNum from R3_ObsCommon..DOC1 where LogNum = " & Session("IORetVal") & " and ItemCode = T0.ItemCode collate database_default and WhsCode = N'" & Session("bodega") & "')), 0) AddedQty " & _
"from OITM T0 " & _
"where T0.ItemCode = N'" & saveHTMLDecode(Request("ItemCode"), False) & "'"
set rs = conn.execute(sql)
ChkSerial = rs("ChkSerial") %>
<script language="javascript">
function onScan(ev){
var scan = ev.data;
	document.frmAddSerNum.txtSerNum.value = scan.value;
	document.frmAddSerNum.submit();
}
function onSwipe(ev){
}
try
{
document.addEventListener("BarcodeScanned", onScan, false);
document.addEventListener("MagCardSwiped", onSwipe, false);
}
catch(err) {}
<% If Request("retVal") <> "" Then %>
<% Select Case Request("retVal")
	Case "OSRIExists" %>alert('<%=getdelOrderCheckSerialLngStr("LtxtOSRIExists")%>'.replace('{0}', '<%=Replace(Request("SuppSerial"), "'", "\'")%>'));
<%	Case "OSRINotFou" %>alert('<%=getdelOrderCheckSerialLngStr("LtxtOSRINotFound")%>'.replace('{0}', '<%=Replace(Request("SuppSerial"), "'", "\'")%>'));
<%	Case "OSRIUsed" %>alert('<%=getdelOrderCheckSerialLngStr("LtxtOSRIUsed")%>'.replace('{0}', '<%=Replace(Request("SuppSerial"), "'", "\'")%>'));
<%	Case "R3Exists" %>alert('<%=getdelOrderCheckSerialLngStr("LtxtR3Exists")%>'.replace('{0}', '<%=Replace(Request("SuppSerial"), "'", "\'")%>'));
<%	Case "SerNotInDo" %>alert('<%=getdelOrderCheckSerialLngStr("LtxtSerNotInDoc")%>'.replace('{0}', '<%=Replace(Request("SuppSerial"), "'", "\'")%>'));
<% End Select %>
<% End If %>

function goBackToCheck()
{
	<% If CDbl(rs("ReqQty")) > CDbl(rs("AddedQty")) Then %>
		<% If ChkSerial = "C" Then %>
		if (!confirm('<%=getdelOrderCheckSerialLngStr("LtxtConfInc")%>')) return;
		<% ElseIf ChkSerial = "E" Then %>
		alert('<%=getdelOrderCheckSerialLngStr("LtxtValInc")%>');
		return;
		<% End If %>
	<% End If %>
	window.location.href='?cmd=<% If Request("retAddPack") <> "Y" Then %>invChkInOutCheck<% Else %>invChkInOutAddByPack<% End If %>&txtOrderNum=<%=Request("txtOrderNum")%><% If Request("retAddPack") = "Y" Then %>&confirm=<%=Server.HTMLEncode(Request("ItemCode"))%><% End If %>';
}
function delSerial(LineNum2, SuppSerial)
{
	if (confirm('<%=getdelOrderCheckSerialLngStr("LtxtConfSerial")%>'.replace('{0}', SuppSerial)))
	{
		window.location.href='inv/delOrderCheckSerialSubmit.asp?txtOrderNum=<%=Request("txtOrderNum")%>&ItemCode=<%=Request("ItemCode")%>&ViewAll=<%=Request("ViewAll")%>&delSerial=' + LineNum2;
	}
}
function focusSerNum()
{
	<% If CDbl(rs("ReqQty")) > CDbl(rs("AddedQty")) Then %>
	document.frmAddSerNum.txtSerNum.focus();
	<% End If %>
}
</script>
<div align="center">
	<center>
	<table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111" bgcolor="#9BC4FF">
		<tr>
			<td bgcolor="#9BC4FF">
			<table border="0" cellpadding="0" bordercolor="#111111" width="100%" id="AutoNumber1">
		        <tr>
		          <td width="100%">
		          <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><!--#include file="delOrderTitle.asp"-->
		          </font></b></td>
		        </tr>
				<tr>
					<td width="100%">
					<p align='<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>'>
					<b><font face="Verdana" size="1"><%=getdelOrderCheckSerialLngStr("LtxtSalesCheckSerial")%>
					</font></b></p>
					</td>
				</tr>
			</table>
			</td>
		</tr>
 		<input type="hidden" name="ItemCode" value="<%=myHTMLEncode(rs("ItemCode"))%>">
		<tr>
			<td bgcolor="#9BC4FF">
			<table cellpadding="0" cellspacing="1" border="0" width="100%">
				<tr>
					<td bgcolor="#66A4FF"><b><font face="Verdana" size="1"><%=getdelOrderCheckSerialLngStr("DtxtItemCode")%>&nbsp;</font></b></td>
					<td bgcolor="#82B4FF"><font face="Verdana" size="1"><%=rs("ItemCode")%></font></td>
				</tr>
				<tr>
					<td bgcolor="#66A4FF"><b><font face="Verdana" size="1"><%=getdelOrderCheckSerialLngStr("LtxtBarCode")%>&nbsp;</font></b></td>
					<td bgcolor="#82B4FF"><font face="Verdana" size="1"><%=rs("CodeBars")%></font></td>
				</tr>
				<tr>
					<td bgcolor="#66A4FF"><b><font face="Verdana" size="1"><%=getdelOrderCheckSerialLngStr("DtxtDescription")%>&nbsp;</font></b></td>
					<td bgcolor="#82B4FF"><font face="Verdana" size="1"><%=rs("ItemName")%></font></td>
				</tr>
				<tr>
					<td bgcolor="#66A4FF"><b><font face="Verdana" size="1"><%=getdelOrderCheckSerialLngStr("LtxtReqQty")%>&nbsp;</font></b></td>
					<td bgcolor="#82B4FF" align="right"><font face="Verdana" size="1"><%=rs("ReqQty")%></font></td>
				</tr>
				<tr>
					<td bgcolor="#66A4FF"><b><font face="Verdana" size="1"><%=getdelOrderCheckSerialLngStr("LtxtProcQty")%>&nbsp;</font></b></td>
					<td bgcolor="#82B4FF" align="right"><font face="Verdana" size="1"><%=rs("AddedQty")%></font></td>
				</tr>
				<tr>
					<td bgcolor="#66A4FF"><b><font face="Verdana" size="1"><%=getdelOrderCheckSerialLngStr("LtxtStatus")%>&nbsp;</font></b></td>
					<td bgcolor="#82B4FF" align="center"><b><font face="Verdana" size="1">
					<% If CDbl(rs("AddedQty")) = 0 Then %><%=getdelOrderCheckSerialLngStr("DtxtPend")%>
					<% ElseIf CDbl(rs("AddedQty")) < CDbl(rs("ReqQty")) Then %><font color="green"><span class="style1">&nbsp;&nbsp;<%=getdelOrderCheckSerialLngStr("LtxtPartial")%>&nbsp;&nbsp;</span></font>
					<% ElseIf CDbl(rs("AddedQty")) = CDbl(rs("ReqQty")) Then %><span class="style2">&nbsp;&nbsp;<%=getdelOrderCheckSerialLngStr("LtxtCompleted")%>&nbsp;&nbsp;</span>
					<% End If %></font></b></td>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td bgcolor="#9BC4FF" height="4"></td>
		</tr>
		<tr>
			<td bgcolor="#9BC4FF" align="center">
			<input type="button" name="btnGoToCheck" value="<%=getdelOrderCheckSerialLngStr("LtxtBackToCheck")%>" onclick="javascript:goBackToCheck()">
			</td>
		</tr>
		<% If CDbl(rs("ReqQty")) > CDbl(rs("AddedQty")) Then %>
		<tr>
			<td bgcolor="#9BC4FF">
			<form method="post" action="inv/delOrderCheckSerialSubmit.asp" name="frmAddSerNum" onsubmit="javascript:return valFrm();">
			<table border="0" cellpadding="0" bordercolor="#111111" width="100%" id="AutoNumber1">
				<tr>
					<td align="center">
					<input type="text" name="txtSerNum" value="" size="30" class="input">
					</td>
				</tr>
				<tr>
					<td align="center">
					<input type="submit" name="btnAdd" value="<%=getdelOrderCheckSerialLngStr("DtxtAdd")%>">
					</td>
				</tr>
			</table>
			<input type="hidden" name="txtOrderNum" value="<%=Request("txtOrderNum")%>">
			<input type="hidden" name="ItemCode" value="<%=Request("ItemCode")%>">
			<input type="hidden" name="retAddPack" value='<%=Request("retAddPack")%>'>
			</form>
			</td>
		</tr>
		<% End If %>
		<% 
		sql = "select Count('') from R3_ObsCommon..DOC4 where Lognum = " & Session("IORetVal") & " and LineNum in (" & _
				"select LineNum from R3_ObsCommon..DOC1 where LogNum = " & Session("IORetVal") & " and ItemCode = N'" & saveHTMLDecode(Request("ItemCode"), False) & "' and WhsCode = N'" & Session("bodega") & "') "
		set rs = conn.execute(sql)
		serialCount = rs(0)
		rs.close
		
		If Request("ViewAll") <> "Y" and serialCount > 10 Then sqlAdd = "top 10" Else sqlAdd = ""
		sql = "select " & sqlAdd & " LineNum2, SuppSerial from R3_ObsCommon..DOC4 where Lognum = " & Session("IORetVal") & " and LineNum in (" & _
				"select LineNum from R3_ObsCommon..DOC1 where LogNum = " & Session("IORetVal") & " and ItemCode = N'" & saveHTMLDecode(Request("ItemCode"), False) & "' and WhsCode = N'" & Session("bodega") & "') " & _
				"order by SysSerial desc"
		set rs = conn.execute(sql) %>
		<tr>
			<td bgcolor="#9BC4FF"><b><font face="Verdana" size="1"><%=getdelOrderCheckSerialLngStr("LtxtProcesedData")%></font></b>
			</td>
		</tr>
		<tr>
			<td bgcolor="#9BC4FF">
			<table cellpadding="0" cellspacing="1" border="0" width="100%">
				<% If Not rs.Eof Then %>
				<% do while not rs.eof %>
				<tr>
					<td bgcolor="#66A4FF"><font face="Verdana" size="1"><%=rs("SuppSerial")%>&nbsp;</font></td>
					<td bgcolor="#66A4FF" align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>" width="16">
					<a href="javascript:delSerial(<%=rs("LineNum2")%>, '<%=Replace(rs("SuppSerial"), "'", "\'")%>');"><img border="0" src="../images/remove.gif"></a></td>
				</tr>
				<% rs.movenext
				loop
				If serialCount > 10 Then %>
				<tr>
					<td colspan="2" bgcolor="#66A4FF" align="center"><input type="button" name="btnViewAll" value="<% If Request("ViewAll") <> "Y" Then %><%=getdelOrderCheckSerialLngStr("LtxtViewAll")%><% Else %><%=getdelOrderCheckSerialLngStr("LtxtViewSummary")%><% End If %>" onclick="javascript:window.location.href='?cmd=invChkInOutCheckSerial&txtOrderNum=<%=Request("txtOrderNum")%>&ItemCode=<%=Request("ItemCode")%>&ViewAll=<% If Request("ViewAll") <> "Y" Then %>Y<% End If %>'"></td>
				</tr>
				<% End If %>
				<% Else %>
				<tr>
					<td colspan="2" bgcolor="#66A4FF" align="center"><font face="Verdana" size="1"><%=getdelOrderCheckSerialLngStr("DtxtNoData")%></font></td>
				</tr>
				<% End If %>
			</table>
			</td>
		</tr>
	</table>
	</center>
</div>
<script language="javascript">
function valFrm()
{
	if (document.frmAddSerNum.txtSerNum.value == '')
	{
		alert('<%=getdelOrderCheckSerialLngStr("LtxtValSerNum")%>');
		document.frmAddSerNum.txtSerNum.focus();
		return false;
	}
	return true;
}
</script>
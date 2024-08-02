
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

<% addLngPathStr = "inv/" %>
<!--#include file="lang/delOrderCheckItemSearch.asp" -->
<% 

sql = "select T1.ChkShowReqSum, T1.ChkAllowOverload, T2.T1 oTable, T2.T2 oTable1, T1.DocType " & _
"from OLKCommon T0 " & _
"inner join OLKInOutSettings T1 on T1.ObjectCode = " & Session("ObjCode") & " and T1.Type = '" & Session("Type") & "' " & _
"inner join OLKDocConf T2 on T2.ObjectCode = T1.ObjectCode "
set rs = conn.execute(sql)
ShowReqSum = rs("ChkShowReqSum") = "Y"
AllowOverload = rs("ChkAllowOverload") = "Y"
oTable = rs("oTable")
oTable1 = rs("oTable1")
If rs("DocType") = "S" Then
	NumIn = "Sale"
	Pack = "Sal"
	UnitMsr = "Sal"
ElseIf rs("DocType") = "P" Then
	NumIn = "Buy"
	Pack = "Pur"
	UnitMsr = "Buy"
End If

If myApp.EnableCodeBarsQry Then
	sql = "declare @CodeBars nvarchar(50) set @CodeBars = N'" & saveHTMLDecode(Request("txtItem"), False) & "' "
	sql = sql & "set @CodeBars = (" & myApp.CodeBarsQry & ") select @CodeBars CodeBars"
	set rs = conn.execute(sql)
	strCodeBars = saveHTMLDecode(rs("CodeBars"), False)
Else
	strCodeBars = saveHTMLDecode(Request("txtItem"), False)
End If
rs.close

sql = 	"select T2.ItemCode, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'ItemName', T2.ItemCode, T2.ItemName) ItemName, " & _
		"OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', '" & Pack & "Packmsr', T2.ItemCode, T2." & Pack & "PackMsr) Packmsr, " & _
		"OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', '" & UnitMsr & "UnitMsr', T2.ItemCode, T2." & UnitMsr & "UnitMsr) UnitMsr, " & _
		"T2.CodeBars, T2." & Pack & "PackUn PackUn, T2.NumIn" & NumIn & " NumIn, T2.PicturName, T2.ManSerNum " & _
		"from " & oTable1 & " T0 " & _
		"inner join " & oTable & " T1 on T1.DocEntry = T0.DocEntry and T1.DocNum = " & Request("txtOrderNum") & " " & _
		"inner join OITM T2 on T2.ItemCode = T0.ItemCode " & _
		"where T0.WhsCode = N'" & Session("bodega") & "' and T0.LineStatus = 'O' " & _  
		"and (T2.ItemCode = N'" & saveHTMLDecode(Request("txtItem"), False) & "' or " & _
		"T2.CodeBars = N'" & strCodeBars & "' "
	
If myApp.EnableSearchItmSupp Then
	sql = sql & " or T2.SuppCatNum = N'" & saveHTMLDecode(Request("txtItem"), False) & "'"
End If

sql = sql & ") Group By T2.ItemCode, T2.ItemName, T2.CodeBars, T2." & Pack & "PackMsr, T2." & Pack & "PackUn, T2." & UnitMsr & "UnitMsr, T2.NumIn" & NumIn & ", T2.PicturName, T2.ManSerNum"

rs.open sql, conn, 3, 1
 %>
<script language="javascript">
<% If Not rs.Eof Then %>
function valFrm()
{
	if (document.frmConfirm.txtUnit.value == '' || document.frmConfirm.txtSBUnit.value == '' || document.frmConfirm.txtPackUnit.value == '')
	{
		alert('<%=getdelOrderCheckItemSearchLngStr("LtxtValUnitValue")%>');
		return false;
	}
	
	if (!MyIsNumeric(document.frmConfirm.txtUnit.value))
	{
		alert('<%=getdelOrderCheckItemSearchLngStr("DtxtValNumVal")%>');
		document.frmConfirm.txtUnit.focus();
		return false;
	}
	
	if (parseInt(document.frmConfirm.txtUnit.value) < 0)
	{
		alert('<%=getdelOrderCheckItemSearchLngStr("DtxtValNumMinVal")%>'.replace('{0}', '0'));
		document.frmConfirm.txtUnit.value = 0;
		document.frmConfirm.txtUnit.focus();
		return false;
	}
	
	if (!MyIsNumeric(document.frmConfirm.txtSBUnit.value))
	{
		alert('<%=getdelOrderCheckItemSearchLngStr("DtxtValNumVal")%>');
		document.frmConfirm.txtSBUnit.focus();
		return false;
	}
	
	if (parseInt(document.frmConfirm.txtSBUnit.value) < 0)
	{
		alert('<%=getdelOrderCheckItemSearchLngStr("DtxtValNumMinVal")%>'.replace('{0}', '0'));
		document.frmConfirm.txtSBUnit.value = 0;
		document.frmConfirm.txtSBUnit.focus();
		return false;		
	}
	
	if (!MyIsNumeric(document.frmConfirm.txtPackUnit.value))
	{
		alert('<%=getdelOrderCheckItemSearchLngStr("DtxtValNumVal")%>');
		document.frmConfirm.txtPackUnit.focus();
		return false;
	}
	
	if (parseInt(document.frmConfirm.txtPackUnit.value) < 0)
	{
		alert('<%=getdelOrderCheckItemSearchLngStr("DtxtValNumMinVal")%>'.replace('{0}', '0'));
		document.frmConfirm.txtPackUnit.value = 0;
		document.frmConfirm.txtPackUnit.focus();
		return false;	
	}
	
	if (parseFloat(document.frmConfirm.txtUnit.value) == 0 && parseFloat(document.frmConfirm.txtSBUnit.value) == 0)
	{
		alert('<% If NumIn = "Sale" Then %><%=getdelOrderCheckItemSearchLngStr("LtxtValSalOrUnit")%><% ElseIf NumIn = "Buy" Then %><%=getdelOrderCheckItemSearchLngStr("LtxtValPurOrUnit")%><% End If %>');
		document.frmConfirm.txtSBUnit.focus();
		return false;
	}
	
	if (parseInt(document.frmConfirm.txtPackUnit.value) == 0 && !document.frmConfirm.chkRepack.checked)
	{
		alert('<%=getdelOrderCheckItemSearchLngStr("LtxtValPackQty")%>');
		document.frmConfirm.txtPackUnit.focus();
		return false;
	}
	
	curUnit = parseFloat(document.frmConfirm.txtUnit.value)+(parseFloat(document.frmConfirm.txtSBUnit.value)*<%=rs("NumIn")%>);
	<% If Not AllowOverload Then %>
	totalReqUnit = reqUnit+(reqSaleUnit*<%=rs("NumIn")%>);
	
	if (curUnit > totalReqUnit)
	{
		alert('<%=getdelOrderCheckItemSearchLngStr("LtxtValCurUnit")%>');
		document.frmConfirm.txtUnit.focus();
		return false;
	}
	/*if(parseFloat(document.frmConfirm.txtUnit.value) > reqUnit)
	{
		alert('<%=getdelOrderCheckItemSearchLngStr("LtxtValReqUnit")%>');
		document.frmConfirm.txtUnit.focus();
		return false;
	}
	
	if (parseFloat(document.frmConfirm.txtSBUnit.value) > reqSaleUnit)
	{
		alert('<% If NumIn = "Sale" Then %><%=getdelOrderCheckItemSearchLngStr("LtxtValReqSaleUnit")%><% ElseIf NumIn = "Buy" Then %><%=getdelOrderCheckItemSearchLngStr("LtxtValReqPurUnit")%><% End If %>');
		document.frmConfirm.txtSBUnit.focus();
		return false;
	}*/
	<% End If %>
	
	if (Math.ceil(curUnit/<%=rs("NumIn")%>/<%=rs("PackUn")%>) != parseFloat(document.frmConfirm.txtPackUnit.value) && !document.frmConfirm.chkRepack.checked)
	{
		alert('<%=getdelOrderCheckItemSearchLngStr("LtxtValCalQuantity")%>');
		document.frmConfirm.chkRepack.focus();
		return false;
	
	}
	
	return true;
}
<% End If %>
</script>
<form name="frmConfirm" method="post" action="inv/delOrderCheckItemSubmit.asp" onsubmit="javascript:return valFrm();">
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
					<b><font face="Verdana" size="1"><%=getdelOrderCheckItemSearchLngStr("LtxtSalesCheckResult")%>
					</font></b></p>
					</td>
				</tr>
			</table>
			</td>
		</tr>
		<% If Not rs.Eof Then
		If rs.recordcount = 1 Then
		If rs("PicturName") <> "" Then
			Pic = rs("PicturName")'
		Else
			Pic = "na.jpg"
		End If 
 		%>
 		<input type="hidden" name="ItemCode" value="<%=myHTMLEncode(rs("ItemCode"))%>">
		<tr>
			<td bgcolor="#9BC4FF">
			<table cellpadding="0" cellspacing="1" border="0" width="100%">
				<tr>
					<td bgcolor="#66A4FF"><b><font face="Verdana" size="1"><%=getdelOrderCheckItemSearchLngStr("DtxtItemCode")%>&nbsp;</font></b></td>
					<td bgcolor="#82B4FF"><font face="Verdana" size="1"><%=rs("ItemCode")%></font></td>
				</tr>
				<tr>
					<td bgcolor="#66A4FF"><b><font face="Verdana" size="1"><%=getdelOrderCheckItemSearchLngStr("LtxtBarCode")%>&nbsp;</font></b></td>
					<td bgcolor="#82B4FF"><font face="Verdana" size="1"><%=rs("CodeBars")%></font></td>
				</tr>
				<tr>
					<td bgcolor="#66A4FF"><b><font face="Verdana" size="1"><%=getdelOrderCheckItemSearchLngStr("DtxtDescription")%>&nbsp;</font></b></td>
					<td bgcolor="#82B4FF"><font face="Verdana" size="1"><%=rs("ItemName")%></font></td>
				</tr>
				<tr>
					<td bgcolor="#82B4FF" colspan="2">
					<table cellpadding="0" cellspacing="0" border="0" width="100%">
					<tr>
						<td colspan="3" height="5px"></td>
					</tr>
					<tr>
						<td rowspan="4" align="center"><a href="delOrderCheckItemSearch.asp?cmd=viewImage&amp;FileName=<%=Pic%>"><img border="0" src="pic.aspx?filename=<%=Pic%>&dbName=<%=Session("olkdb")%>&MaxSize=80"></a></td>
						<td><font face="Verdana" size="1"><%=getdelOrderCheckItemSearchLngStr("DtxtBaseUnit")%></font>
						</td>
						<td>
						<input type="number" size="6" name="txtUnit" <% If CDbl(rs("NumIn")) = 1 Then %>disabled<% End If %> value="0" style="text-align: right; height: 22px;"></td>
					</tr>
					<tr>
						<td><font face="Verdana" size="1"><% If NumIn = "Sale" Then %><%=getdelOrderCheckItemSearchLngStr("DtxtSalUnit")%><% ElseIf NumIn = "Buy" Then %><%=getdelOrderCheckItemSearchLngStr("DtxtBuyUnit")%><% End If %></font>
						</td>
						<td>
						<input type="number" name="txtSBUnit" size="6" value="0" style="text-align: right" min="0" step="<%=GetNumberStep(myApp.QtyDec)%>"></td>
					</tr>
					<tr>
						<td><font face="Verdana" size="1"><%=getdelOrderCheckItemSearchLngStr("DtxtPackUnit")%></font>
						</td>
						<td>
						<input type="number" name="txtPackUnit" size="6" value="0" style="text-align: right" min="0" step="<%=GetNumberStep(myApp.QtyDec)%>"></td>
					</tr>
					<tr>
						<td colspan="2">
						<input type="checkbox" name="Repack" id="chkRepack" value="Y" style="height: 20px"><font face="Verdana" size="1"><label for="chkRepack"><%=getdelOrderCheckItemSearchLngStr("LtxtRepack")%></label></font></td>
					</tr>
					<tr>
						<td colspan="3" height="5px"></td>
					</tr>
					<% If rs("ManSerNum") = "Y" Then %>
					<tr>
						<td colspan="3" align="center"><input type="button" name="btnSerial" value="<%=getdelOrderCheckItemSearchLngStr("LtxtSerialNumbers")%>" onclick="javascript:window.location.href='operaciones.asp?cmd=invChkInOutCheckSerial&txtOrderNum=<%=Request("txtOrderNum")%>&ItemCode=<%=rs("ItemCode")%>';"></td>
					</tr>
					<tr>
						<td colspan="3" height="5px"></td>
					</tr>
					<% End If %>
					<tr>
						<td align="center" colspan="3">
						<table cellpadding="0" cellspacing="0" border="0" width="100%">
							<tr>
								<td align="center" width="50%"><input type="button" name="btnClear" value="<%=getdelOrderCheckItemSearchLngStr("DtxtClear")%>" onclick="javascript:confirmClear('<%=Replace(rs("ItemCode"), "'", "\'")%>');">
								</td>
								<td align="center" width="50%"><input type="submit" name="btnConfirm" value="<%=getdelOrderCheckItemSearchLngStr("DtxtConfirm")%>"></td>
							</tr>
						</table>
						</td>
					</tr>
					<tr>
						<td colspan="3" height="5px"></td>
					</tr>
				</table></td>
				</tr>
				<% set rp = Server.CreateObject("ADODB.RecordSet")
				sql = "declare @ItemCode nvarchar(20) set @ItemCode = N'" & saveHTMLDecode(rs("ItemCode"), false) & "' " & _
				"declare @DocEntry int set @DocEntry = (select DocEntry from " & oTable & " where DocNum = " & Request("txtOrderNum") & ") " & _
				"declare @LogNum int set @LogNum = " & Session("IORetVal") & " " & _
				"declare @Quantity numeric(19,6)  " & _
				"declare @CompareQty numeric(19,6) " & _
				"set @Quantity = ( " & _
				"select Sum(Case UseBaseUn When 'Y' Then T0.OpenQty Else T0.OpenQty*T1.NumIn" & NumIn & " End) " & _
				"from " & oTable1 & " T0 " & _
				"inner join OITM T1 on T1.ItemCode = T0.ItemCode " & _
				"where T0.DocEntry = @DocEntry and T0.ItemCode = @ItemCode) " & _
				"set @CompareQty = IsNull(( " & _
				"select Sum(Case UseBaseUn When 'Y' Then T0.Quantity Else T0.Quantity*T1.NumIn" & NumIn & " End) " & _
				"from R3_ObsCommon..DOC1 T0 " & _
				"inner join OITM T1 on T1.ItemCode = T0.ItemCode collate database_default " & _
				"where T0.LogNum = @LogNum and T0.ItemCode = @ItemCode), 0) " & _
				"select case when @CompareQty = 0 Then 'O' " & _
				"	When @CompareQty < @Quantity Then 'P' " & _
				"	When @CompareQty = @Quantity Then 'C' " & _
				"	When @CompareQty > @Quantity Then 'E' " & _
				"End Verfy, @Quantity, @CompareQty "
				set rp = conn.execute(sql) %>
				<tr>
					<td bgcolor="#66A4FF"><b><font face="Verdana" size="1"><%=getdelOrderCheckItemSearchLngStr("LtxtStatus")%>&nbsp;</font></b></td>
					<td bgcolor="#82B4FF" class="style5"><b><font face="Verdana" size="1"><% Select Case rp("Verfy")
						Case "O" %><%=getdelOrderCheckItemSearchLngStr("DtxtPend")%>
						<% Case "P" %><font color="green"><span class="style1">&nbsp;&nbsp;<%=getdelOrderCheckItemSearchLngStr("LtxtPartial")%>&nbsp;&nbsp;</span></font>
						<% Case "C" %><span class="style2">&nbsp;&nbsp;<%=getdelOrderCheckItemSearchLngStr("LtxtCompleted")%>&nbsp;&nbsp;</span>
						<% Case "E" %><font color="#CC3300"><%=getdelOrderCheckItemSearchLngStr("LtxtOverload")%></font>
						<% End Select %></font></b></td>
				</tr>
				<% sql = "select IsNull(Sum(Case T0.UseBaseUn When 'N' Then 0 When 'Y' Then T0.Quantity End), 0) Unit, " & _
							"		IsNull(Sum(Case T0.UseBaseUn When 'N' Then T0.Quantity When 'Y' Then 0 End), 0) " & NumIn & "Unit, " & _
							"		IsNull(Sum(T0.PackQty), 0) PackUnit " & _
							"from R3_ObsCommon..DOC1 T0 " & _
							"where T0.LogNum = " & Session("IORetVal") & " and ItemCode = N'" & saveHTMLDecode(rs("ItemCode"), false) & "' and T0.WhsCode = N'" & Session("bodega") & "' "
					set rp = conn.execute(sql) %>
				<tr>
					<td bgcolor="#9BC4FF" colspan="2">
					<table cellpadding="0" cellspacing="1" border="0" width="100%">
						<tr>
							<td bgcolor="#66A4FF" colspan="2" class="style5"><font face="Verdana" size="1">&nbsp;</font></td>
							<td bgcolor="#66A4FF" class="style5"><b><font face="Verdana" size="1"><%=getdelOrderCheckItemSearchLngStr("DtxtUnit")%></font></b></td>
							<td bgcolor="#66A4FF" class="style5"><b><font face="Verdana" size="1"><%=rs("UnitMsr")%>&nbsp;(<%=rs("NumIn")%>)</font></b></td>
							<td bgcolor="#66A4FF" class="style5"><b><font face="Verdana" size="1"><%=rs("Packmsr")%>&nbsp;(<%=rs("PackUn")%>)</font></b></td>
						</tr>
						<tr>
							<td bgcolor="#66A4FF" colspan="2" style="height: 5px"><b><font face="Verdana" size="1"><%=getdelOrderCheckItemSearchLngStr("LtxtCurSum")%>&nbsp;</font></b></td>
							<td bgcolor="#82B4FF" class="style5" style="height: 5px"><font face="Verdana" size="1"><%=FormatNumber(CDbl(rp("Unit")), myApp.QtyDec)%></font></td>
							<td bgcolor="#82B4FF" class="style5" style="height: 5px"><font face="Verdana" size="1"><%=FormatNumber(CDbl(rp(NumIn & "Unit")), myApp.QtyDec)%></font></td>
							<td bgcolor="#82B4FF" class="style5" style="height: 5px"><font face="Verdana" size="1"><%=FormatNumber(CDbl(rp("PackUnit")), myApp.QtyDec)%></font></td>
						</tr>
				<% sql = "select Sum(Case UseBaseUn When 'Y' Then T1.OpenQty Else 0 End) Unit, " & _
						"Sum(Case UseBaseUn When 'N' Then T1.OpenQty Else 0 End) " & NumIn & "Unit, " & _
						"Ceiling(Sum(Case UseBaseUn When 'N' Then T1.OpenQty Else 0 End)/" & Pack & "PackUn) PackUnit " & _
						"from " & oTable & " T0 " & _
						"inner join " & oTable1 & " T1 on T1.DocEntry = T0.DocEntry " & _
						"inner join OITM T2 on T2.ItemCode = T1.ItemCode " & _
						"where T0.DocNum = " & Request("txtOrderNum") & " and T1.ItemCode = N'" & saveHTMLDecode(rs("ItemCode"), false) & "' " & _
						"Group By T2." & Pack & "PackUn"
					set rp = conn.execute(sql)
					If ShowReqSum Then %>
						<tr>
							<td bgcolor="#66A4FF" colspan="2"><b><font face="Verdana" size="1"><%=getdelOrderCheckItemSearchLngStr("LtxtReqSum")%>&nbsp;</font></b></td>
							<td bgcolor="#82B4FF" class="style5"><font face="Verdana" size="1"><%=FormatNumber(CDbl(rp("Unit")), myApp.QtyDec)%></font></td>
							<td bgcolor="#82B4FF" class="style5"><font face="Verdana" size="1"><%=FormatNumber(CDbl(rp(NumIn & "Unit")), myApp.QtyDec)%></font></td>
							<td bgcolor="#82B4FF" class="style5"><font face="Verdana" size="1"><%=FormatNumber(CDbl(rp("PackUnit")), myApp.QtyDec)%></font></td>
						</tr>
					</table>
					<% End If %>
					<script language="javascript">
					var reqUnit = <%=rp("Unit")%>;
					var reqSaleUnit = <%=rp(NumIn & "Unit")%>;
					</script></td>
				</tr>
			</table>
			</td>
		</tr>		
		<% Else
		sql = "select T0.LineNum+1 LineNum, T0.ItemCode, T2.CodeBars " & _
		"from " & oTable1 & " T0 " & _
		"inner join " & oTable & " T1 on T1.DocEntry = T0.DocEntry and T1.DocNum = " & Request("txtOrderNum") & " " & _
		"inner join OITM T2 on T2.ItemCode = T0.ItemCode " & _
		"where T0.LineStatus = 'O' and T0.WhsCode = N'" & Session("bodega") & "' and (T2.ItemCode = N'" & saveHTMLDecode(Request("txtItem"), False) & "' or T2.CodeBars = N'" & strCodeBars & "') " & _
		"order by LineNum"
		rs.close
		rs.open sql, conn, 3, 1 %>
		<tr>
			<td bgcolor="#9BC4FF" style="text-align: justify; "><font face="Verdana" size="1"><%=Replace(getdelOrderCheckItemSearchLngStr("LtxtConflictFound"), "{0}", rs.recordcount)%></font></td>
		</tr>
		<tr>
			<td bgcolor="#9BC4FF" height="5px; "></td>
		</tr>
		<tr>
			<td bgcolor="#9BC4FF">
			<table cellpadding="0" cellspacing="1" border="0" width="100%">
				<tr>
					<td bgcolor="#66A4FF"><b><font face="Verdana" size="1">#</font></b></td>
					<td bgcolor="#66A4FF"><b><font face="Verdana" size="1">&nbsp;<%=getdelOrderCheckItemSearchLngStr("DtxtItemCode")%></font></b></td>
					<td bgcolor="#66A4FF"><b><font face="Verdana" size="1">&nbsp;<%=getdelOrderCheckItemSearchLngStr("LtxtBarCode")%></font></b></td>
				</tr>
				<% do while not rs.eof %>
				<tr>
					<td bgcolor="#82B4FF"><font face="Verdana" size="1"><%=rs("LineNum")%></font></td>
					<td bgcolor="#82B4FF"><font face="Verdana" size="1">&nbsp;<%=rs("ItemCode")%></font></td>
					<td bgcolor="#82B4FF"><font face="Verdana" size="1">&nbsp;<%=rs("CodeBars")%></font></td>
				</tr>
				<% rs.movenext
				loop %>
			</table>
			</td>
		</tr>
		<% End If %>
		<% Else
		sql = "select T0.ItemCode, T0.LineNum+1 LineNum " & _
		"from " & oTable1 & " T0 " & _
		"inner join " & oTable & " T1 on T1.DocEntry = T0.DocEntry and T1.DocNum = " & Request("txtOrderNum") & " " & _
		"inner join OITM T2 on T2.ItemCode = T0.ItemCode " & _
		"where T0.WhsCode = N'" & Session("bodega") & "' and T0.LineStatus = 'C' and (T2.ItemCode = N'" & saveHTMLDecode(Request("txtItem"), False) & "' or T2.CodeBars = N'" & strCodeBars & "') "
		set rs = conn.execute(sql)
		If rs.eof Then %>
        <tr>
          <td width="100%" style="text-align: justify">
			<font face="Verdana" size="1"><% If Request("cmd") = "searchDelCheckItem" Then %><%=Replace(Replace(getdelOrderCheckItemSearchLngStr("LtxtNoItmFoundInDoc"), "{0}", Request("txtItem")), "{1}", Request("txtOrderNum"))%><% ElseIf Request("cmd") = "searchPurCheckItem" Then %><%=Replace(Replace(getdelOrderCheckItemSearchLngStr("LtxtNoItmFoundInPur"), "{0}", Request("txtItem")), "{1}", Request("txtOrderNum"))%><% End If %></font></td>
        </tr>
        <% Else %>
        <tr>
          <td width="100%" style="text-align: justify">
			<font face="Verdana" size="1"><% If Request("cmd") = "searchDelCheckItem" Then %><%=Replace(Replace(Replace(getdelOrderCheckItemSearchLngStr("LtxtClosedLineFound"), "{0}", Request("txtItem")), "{1}", Request("txtOrderNum")), "{2}", rs("LineNum"))%><% ElseIf Request("cmd") = "searchPurCheckItem" Then %><%=Replace(Replace(Replace(getdelOrderCheckItemSearchLngStr("LtxtPurClosedLineFoun"), "{0}", Request("txtItem")), "{1}", Request("txtOrderNum")), "{2}", rs("LineNum"))%><% End If %></font></td>
        </tr>
        <% End If %>
        <% End If %>
		<tr>
			<td bgcolor="#9BC4FF" height="5px"></td>
		</tr>
		<tr>
			<td bgcolor="#9BC4FF" align="center"><input type="button" name="btnReturn" value="<%=getdelOrderCheckItemSearchLngStr("DtxtBack")%>" onclick="javascript:window.location.href='?cmd=invChkInOutCheck&txtOrderNum=<%=Request("txtOrderNum")%>'"></td>
		</tr>
	</table>
	</center>
</div>
<input type="hidden" name="saveCmd" value="invChkInOut">
<input type="hidden" name="txtOrderNum" value='<%=Request("txtOrderNum")%>'>
<% If Not rs.Eof Then %>
<input type="hidden" name="NumIn" value='<%=rs("NumIn")%>'>
<% End If %>
</form>
<script language="javascript">
function confirmClear(ItemCode)
{
	document.frmClear.ItemCode.value = ItemCode;
	if (confirm('<%=getdelOrderCheckItemSearchLngStr("LtxtConfClear")%>'))
	{
		document.frmClear.submit();
	}
}
</script>
<form name="frmClear" method="post" action="inv/delOrderCheckItemSubmit.asp">
<input type="hidden" name="txtItem" value="<%=Request("txtItem")%>">
<input type="hidden" name="ItemCode" value="">
<input type="hidden" name="txtOrderNum" value="<%=Request("txtOrderNum")%>">
<input type="hidden" name="saveCmd" value="invChkInOut">
<input type="hidden" name="cmd" value="clear">
</form>
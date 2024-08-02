<!--#include file="top.asp" -->
<!--#include file="lang/adminCustDec.asp" -->

<head>
<style type="text/css">
.style1 {
	text-align: center;
}
</style>
</head>

<script language="javascript" src="js_up_down.js"></script>
<%
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetDecSettingsData" & Session("ID")
cmd.Parameters.Refresh()
set rs = cmd.execute() 
%>
<form method="POST" action="adminsubmit.asp" name="Form1" onsubmit="javascript:return valFrm();">
	<table border="0" cellpadding="0" width="100%">
		<tr>
			<td bgcolor="#E1F3FD"><b><font color="#FFFFFF">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><%=getadminCustDecLngStr("LttlCustDec")%></font></b></td>
		</tr>
		<tr>
			<td bgcolor="#F5FBFE">
			<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1">
			</font><font face="Verdana" size="1" color="#4783C5"><%=getadminCustDecLngStr("LttlCustDecNote")%></font></p>
			</td>
		</tr>
		<tr>
			<td>
				<table border="0" cellpadding="0" width="100%">
					<tr>
						<td bgcolor="#E1F3FD" style="width: 120"></td>
						<td bgcolor="#E1F3FD" class="style1" style="width: 120"><font face="Verdana" size="1" color="#31659C">
						<strong><%=getadminCustDecLngStr("DtxtSystem")%></strong></font></td>
						<td bgcolor="#E1F3FD" class="style1" style="width: 120"><font face="Verdana" size="1" color="#31659C">
						<strong><%=getadminCustDecLngStr("DtxtAlternative")%></strong></font></td>
						<td bgcolor="#E1F3FD">&nbsp;</td>
					</tr>
					<tr>
						<td bgcolor="#E1F3FD" style="width: 120"><font face="Verdana" size="1" color="#31659C"><%=getadminCustDecLngStr("DtxtQty")%></font></td>
						<td bgcolor="#F7FBFF" style="width: 120" class="style1"><font face="Verdana" size="1" color="#4783C5"><%=rs("QtyDec")%></font></td>
						<td bgcolor="#F7FBFF" style="width: 120" class="style1">
						<input type="text" name="AlterQtyDec" id="AlterQtyDec" size="1" value='<%=rs("AlterQtyDec")%>' class="input" onfocus="this.select()" maxlength="1" onkeydown="return valKeyNum(event);" style="text-align: right;">
						</td>
						<td bgcolor="#E1F3FD">&nbsp;</td>
					</tr>
					<tr>
						<td bgcolor="#E1F3FD" style="width: 120"><font face="Verdana" size="1" color="#31659C"><%=getadminCustDecLngStr("DtxtPrice")%></font></td>
						<td bgcolor="#F7FBFF" style="width: 120" class="style1"><font face="Verdana" size="1" color="#4783C5"><%=rs("PriceDec")%></font></td>
						<td bgcolor="#F7FBFF" style="width: 120" class="style1">
						<input type="text" name="AlterPriceDec" id="AlterPriceDec" size="1" value='<%=rs("AlterPriceDec")%>' class="input" onfocus="this.select()" maxlength="1" onkeydown="return valKeyNum(event);" style="text-align: right;"></td>
						<td bgcolor="#E1F3FD">&nbsp;</td>
					</tr>
					<tr>
						<td bgcolor="#E1F3FD" style="width: 120"><font face="Verdana" size="1" color="#31659C"><%=getadminCustDecLngStr("DtxtPercentage")%></font></td>
						<td bgcolor="#F7FBFF" style="width: 120" class="style1"><font face="Verdana" size="1" color="#4783C5"><%=rs("PercentDec")%></font></td>
						<td bgcolor="#F7FBFF" style="width: 120" class="style1">
						<input type="text" name="AlterPercentDec" id="AlterPercentDec" size="1" value='<%=rs("AlterPercentDec")%>' class="input" onfocus="this.select()" maxlength="1" onkeydown="return valKeyNum(event);" style="text-align: right;"></td>
						<td bgcolor="#E1F3FD">&nbsp;</td>
					</tr>
					<tr>
						<td bgcolor="#E1F3FD" style="width: 120"><font face="Verdana" size="1" color="#31659C"><%=getadminCustDecLngStr("DtxtMeasure")%></font></td>
						<td bgcolor="#F7FBFF" style="width: 120" class="style1"><font face="Verdana" size="1" color="#4783C5"><%=rs("MeasureDec")%></font></td>
						<td bgcolor="#F7FBFF" style="width: 120" class="style1">
						<input type="text" name="AlterMeasureDec" id="AlterMeasureDec" size="1" value='<%=rs("AlterMeasureDec")%>' class="input" onfocus="this.select()" maxlength="1" onkeydown="return valKeyNum(event);" style="text-align: right;"></td>
						<td bgcolor="#E1F3FD">&nbsp;</td>
					</tr>
					<tr>
						<td bgcolor="#E1F3FD" style="width: 120"><font face="Verdana" size="1" color="#31659C"><%=getadminCustDecLngStr("DtxtSum")%></font></td>
						<td bgcolor="#F7FBFF" style="width: 120" class="style1"><font face="Verdana" size="1" color="#4783C5"><%=rs("SumDec")%></font></td>
						<td bgcolor="#F7FBFF" style="width: 120" class="style1">
						<input type="text" name="AlterSumDec" id="AlterSumDec" size="1" value='<%=rs("AlterSumDec")%>' class="input" onfocus="this.select()" maxlength="1" onkeydown="return valKeyNum(event);" style="text-align: right;"></td>
						<td bgcolor="#E1F3FD">&nbsp;</td>
					</tr>
					<tr>
						<td bgcolor="#E1F3FD" style="width: 120"><font face="Verdana" size="1" color="#31659C"><%=getadminCustDecLngStr("DtxtRate")%></font></td>
						<td bgcolor="#F7FBFF" style="width: 120" class="style1"><font face="Verdana" size="1" color="#4783C5"><%=rs("RateDec")%></font></td>
						<td bgcolor="#F7FBFF" style="width: 120" class="style1">
						<input type="text" name="AlterRateDec" id="AlterRateDec" size="1" value='<%=rs("AlterRateDec")%>' class="input" onfocus="this.select()" maxlength="1" onkeydown="return valKeyNum(event);" style="text-align: right;"></td>
						<td bgcolor="#E1F3FD">&nbsp;</td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td>
			<table border="0" cellpadding="0" width="100%">
				<tr>
					<td width="77">
					<input type="submit" value="<%=getadminCustDecLngStr("DtxtSave")%>" name="btnSave" class="OlkBtn"></td>
					<td><hr color="#0D85C6" size="1"></td>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td>&nbsp;</td>
		</tr>
	</table>
	<input type="hidden" name="submitCmd" value="adminDecimals">
</form>
<script type="text/javascript">
function valKeyNum(e)
{
	switch (e.keyCode)
	{
		case 8:
		case 9:
		case 16:
		case 38:
		case 36:
		case 46:
		case 48:
		case 49:
		case 50:
		case 51:
		case 52:
		case 53:
		case 54:
		case 96:
		case 97:
		case 98:
		case 99:
		case 100:
		case 101:
		case 102:
		case 37: // Left
		case 39: //Right
			return true;
	}
	return false;
}

</script>
<!--#include file="bottom.asp" -->
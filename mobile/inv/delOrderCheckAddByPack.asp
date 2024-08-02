
<head>
<style type="text/css">
.style1 {
				text-align: center;
}
.style2 {
	background-color: #008000;
	color: #FFFF00;
}
.style3 {
				background-color: #FFFF00;
}
</style>
</head>

<% addLngPathStr = "inv/" %>
<!--#include file="lang/delOrderCheckAddByPack.asp" -->
<%

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "DBOLKGetIOCheckData" & Session("ID")
cmd.Parameters.Refresh()
cmd("@LanID") = Session("LanID")
cmd("@ObjectCode") = Session("ObjCode")
cmd("@Type") = Session("Type")
cmd("@DocNum") = Request("txtOrderNum")
cmd("@LogNum") = Session("IORetVal")
cmd("@WhsCode") = Session("bodega")

set rs = cmd.execute()
Status = rs("Status")
SerialStatus = rs("SerialStatus")
EnablePartial = rs("PartDelivr") = "Y"
AllowOverload = rs("ChkAllowOverload") = "Y"
BackOrder = rs("BackOrder") = "Y"
ObjectCode = rs("ChkOp")
ChkSerial = rs("ChkSerial")
HasSerial = rs("HasSerial") = "Y"
ObjDesc = rs("ObjDesc")
DocNum = Request("txtOrderNum")

%>
<div align="center">
	<center>
	<table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
		<tr>
			<td bgcolor="#9BC4FF">
			<table border="0" cellpadding="0" bordercolor="#111111" width="100%">
			<form method="post" action="inv/delOrderCheckAddByPackSubmit.asp" name="frmAddByPack" onsubmit="javascript:return valFrm();">
		        <tr>
		          <td width="100%">
		          <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><!--#include file="delOrderTitle.asp"-->
		          </font></b></td>
		        </tr>
				<tr>
					<td width="100%">
					<p align='<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>'>
					<b><font face="Verdana" size="1">
					<% 
					varText = getdelOrderCheckAddByPackLngStr("LtxtDocCheck")
					varText = Replace(varText, "{0}", ObjDesc)
					varText = Replace(varText, "{1}", DocNum)
					Response.WRite varText %>
					</font></b></p>
					</td>
				</tr>
				<tr>
					<td width="100%" align="center">
					<table cellpadding="0" cellspacing="1" border="0" width="100%">
						<tr>
							<td bgcolor="#66A4FF"><font size="1" face="Verdana"><b><% If rs("DocType") = "Sale" Then %><%=getdelOrderCheckAddByPackLngStr("DtxtClient")%><% ElseIf rs("DocType") = "Buy" Then %><%=getdelOrderCheckAddByPackLngStr("DtxtProvider")%><% End If %></b></font></td>
							<td bgcolor="#82B4FF"><font size="1" face="Verdana"><%=rs("CardCode")%></font></td>
						</tr>
						<tr>
							<td bgcolor="#66A4FF"><font size="1" face="Verdana"><b><%=getdelOrderCheckAddByPackLngStr("DtxtName")%></b></font></td>
							<td bgcolor="#82B4FF"><font size="1" face="Verdana"><%=rs("CardName")%></font></td>
						</tr>
					</table>
					</td>
				</tr>
				<tr>
					<td width="100%" height="5px"></td>
				</tr>
				<tr>
					<td width="100%">
					<table border="0" cellpadding="0" cellspacing="1" bordercolor="#111111" width="100%" id="AutoNumber2">
						<tr>
							<td>
								<table border="0" cellpadding="0" cellspacing="1" width="100%" bordercolor="#111111">
									<tr>
										<td bgcolor="#66A4FF"><font size="1" face="Verdana"><b><%=getdelOrderCheckAddByPackLngStr("DtxtQty")%></b></font></td>
										<td align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
										<input name="txtQty" type="text" maxlength="25" value="1" size="6" style="text-align: right;"></td>
									</tr>
									<tr>
										<td colspan="2" align="center" bgcolor="#66A4FF"><font size="1" face="Verdana"><b>
										<%=getdelOrderCheckAddByPackLngStr("LtxtItemCodeOrCodBar")%></b></font></td>
									</tr>
									<tr>
										<td colspan="2" align="center">
										<input maxlength="20" name="txtItem" type="text" size="20"></td>
									</tr>
									<tr>
										<td>
										<table>
														<tr bgcolor="#66A4FF">
																		<td width="15"><input name="rdUnit" id="rdUnit3" type="radio" value="3" checked></td>
																		<td><label for="rdUnit3"><font size="1" face="Verdana"><b><%=getdelOrderCheckAddByPackLngStr("DtxtPackUnit")%></b></font></label></td>
														</tr>
														<tr bgcolor="#66A4FF">
																		<td width="15"><input name="rdUnit" id="rdUnit2" type="radio" value="2"></td>
																		<td><label for="rdUnit2"><font size="1" face="Verdana"><b><%=getdelOrderCheckAddByPackLngStr("DtxtUnit")%></b></font></label></td>
														</tr>
														<tr bgcolor="#66A4FF">
																		<td width="15"><input name="rdUnit" id="rdUnit1" type="radio" value="1"></td>
																		<td><label for="rdUnit1"><font size="1" face="Verdana"><b><%=getdelOrderCheckAddByPackLngStr("DtxtBaseUnit")%></b></font></label></td>
														</tr>
										</table>
										</td>
										<td valign="bottom">
										<table cellpadding="0" cellspacing="2" border="0" width="100%">
											<tr>
												<td width="15"><input type="checkbox" name="chkRepack" id="chkRepack" value="Y"></td>
												<td><label for="chkRepack"><font size="1" face="Verdana"><%=getdelOrderCheckAddByPackLngStr("LtxtCreatePackage")%></font></label></td>
											</tr>
											<tr>
												<td colspan="2">&nbsp;</td>
											</tr>
											<tr>
												<td colspan="2" align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><input type="submit" name="btnAdd" value="<%=getdelOrderCheckAddByPackLngStr("DtxtAdd")%>"></td>
											</tr>
										</table>
										</td>
									</tr>
								</table>
								<% If Request("confirm") <> "" Then
								ItemCode = Request("confirm")
								
								sql = "select T1.ChkShowReqSum, T2.T1 oTable, T2.T2 oTable1, T1.DocType " & _
								"from OLKCommon T0 " & _
								"inner join OLKInOutSettings T1 on T1.ObjectCode = " & Session("ObjCode") & " and T1.Type = '" & Session("Type") & "' " & _
								"inner join OLKDocConf T2 on T2.ObjectCode = T1.ObjectCode "
								set rs = conn.execute(sql)
								ShowReqSum = rs("ChkShowReqSum") = "Y"
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
								

								sql = 	"select T2.ItemCode, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'ItemName', T2.ItemCode, T2.ItemName) ItemName, " & _
										"OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', '" & Pack & "Packmsr', T2.ItemCode, T2." & Pack & "PackMsr) Packmsr, " & _
										"OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', '" & UnitMsr & "UnitMsr', T2.ItemCode, T2." & UnitMsr & "UnitMsr) UnitMsr, " & _
										"T2.CodeBars, T2." & Pack & "PackUn PackUn, T2.NumIn" & NumIn & " NumIn, T2.PicturName, T2.ManSerNum " & _
										"from " & oTable1 & " T0 " & _
										"inner join " & oTable & " T1 on T1.DocEntry = T0.DocEntry and T1.DocNum = " & Request("txtOrderNum") & " " & _
										"inner join OITM T2 on T2.ItemCode = T0.ItemCode " & _
										"where T0.WhsCode = N'" & Session("bodega") & "' and T0.LineStatus = 'O' and T2.ItemCode = N'" & saveHTMLDecode(ItemCode, False) & "' " & _
										"Group By T2.ItemCode, T2.ItemName, T2.CodeBars, T2." & Pack & "PackMsr, T2." & Pack & "PackUn, T2." & UnitMsr & "UnitMsr, T2.NumIn" & NumIn & ", T2.PicturName, T2.ManSerNum"
								set rs = conn.execute(sql)
								 %>
								<br>
								<table border="0" cellpadding="0" cellspacing="1" width="100%" bordercolor="#111111">
									<tr>
										<td colspan="2" align="center" bgcolor="#66A4FF"><font size="1" face="Verdana"><b><%=getdelOrderCheckAddByPackLngStr("LtxtLastItem")%>: <%=ItemCode%></b></font></td>
									</tr>

									<tr>
										<td bgcolor="#9BC4FF">
										<table cellpadding="0" cellspacing="1" border="0" width="100%">
											<tr>
												<td bgcolor="#66A4FF"><b><font face="Verdana" size="1"><%=getdelOrderCheckAddByPackLngStr("LtxtBarCode")%>&nbsp;</font></b></td>
												<td bgcolor="#82B4FF"><font face="Verdana" size="1"><%=rs("CodeBars")%></font></td>
											</tr>
											<tr>
												<td bgcolor="#66A4FF"><b><font face="Verdana" size="1"><%=getdelOrderCheckAddByPackLngStr("DtxtDescription")%>&nbsp;</font></b></td>
												<td bgcolor="#82B4FF"><font face="Verdana" size="1"><%=rs("ItemName")%></font></td>
											</tr>
										</table>
										</td>
									</tr>
									<tr>
										<td>
										<table cellpadding="0" cellspacing="1" border="0" width="100%">
												<% set rp = Server.CreateObject("ADODB.RecordSet")
												sql = "declare @ItemCode nvarchar(20) set @ItemCode = N'" & saveHTMLDecode(ItemCode, false) & "' " & _
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
													<td bgcolor="#66A4FF"><b><font face="Verdana" size="1"><%=getdelOrderCheckAddByPackLngStr("LtxtStatus")%>&nbsp;</font></b></td>
													<td bgcolor="#82B4FF" class="style5"><b><font face="Verdana" size="1"><% Select Case rp("Verfy")
														Case "O" %><%=getdelOrderCheckAddByPackLngStr("DtxtPend")%>
														<% Case "P" %><font color="green"><span class="style3 ">&nbsp;&nbsp;<%=getdelOrderCheckAddByPackLngStr("LtxtPartial")%>&nbsp;&nbsp;</span></font>
														<% Case "C" %><span class="style2">&nbsp;&nbsp;<%=getdelOrderCheckAddByPackLngStr("LtxtCompleted")%>&nbsp;&nbsp;</span>
														<% Case "E" %><font color="#CC3300"><%=getdelOrderCheckAddByPackLngStr("LtxtOverload")%></font>
														<% End Select %></font></b></td>
												</tr>
												<% sql = "select IsNull(Sum(Case T0.UseBaseUn When 'N' Then 0 When 'Y' Then T0.Quantity End), 0) Unit, " & _
															"		IsNull(Sum(Case T0.UseBaseUn When 'N' Then T0.Quantity When 'Y' Then 0 End), 0) " & NumIn & "Unit, " & _
															"		IsNull(Sum(T0.PackQty), 0) PackUnit " & _
															"from R3_ObsCommon..DOC1 T0 " & _
															"where T0.LogNum = " & Session("IORetVal") & " and ItemCode = N'" & saveHTMLDecode(ItemCode, false) & "' and T0.WhsCode = N'" & Session("bodega") & "' "
													set rp = conn.execute(sql) %>
												<tr>
													<td bgcolor="#9BC4FF" colspan="2">
													<table cellpadding="0" cellspacing="1" border="0" width="100%">
														<tr>
															<td bgcolor="#66A4FF" colspan="2" class="style5"><font face="Verdana" size="1">&nbsp;</font></td>
															<td bgcolor="#66A4FF" class="style5"><b><font face="Verdana" size="1"><%=getdelOrderCheckAddByPackLngStr("DtxtUnit")%></font></b></td>
															<td bgcolor="#66A4FF" class="style5"><b><font face="Verdana" size="1"><%=rs("UnitMsr")%>&nbsp;(<%=rs("NumIn")%>)</font></b></td>
															<td bgcolor="#66A4FF" class="style5"><b><font face="Verdana" size="1"><%=rs("Packmsr")%>&nbsp;(<%=rs("PackUn")%>)</font></b></td>
														</tr>
														<tr>
															<td bgcolor="#66A4FF" colspan="2"><b><font face="Verdana" size="1"><%=getdelOrderCheckAddByPackLngStr("LtxtCurSum")%>&nbsp;</font></b></td>
															<td bgcolor="#82B4FF" class="style1"><font face="Verdana" size="1"><%=FormatNumber(CDbl(rp("Unit")), myApp.QtyDec)%></font></td>
															<td bgcolor="#82B4FF" class="style1"><font face="Verdana" size="1"><%=FormatNumber(CDbl(rp(NumIn & "Unit")), myApp.QtyDec)%></font></td>
															<td bgcolor="#82B4FF" class="style1"><font face="Verdana" size="1"><%=FormatNumber(CDbl(rp("PackUnit")), myApp.QtyDec)%></font></td>
														</tr>
												<% sql = "select Sum(Case UseBaseUn When 'Y' Then T1.OpenQty Else 0 End) Unit, " & _
														"Sum(Case UseBaseUn When 'N' Then T1.OpenQty Else 0 End) " & NumIn & "Unit, " & _
														"Ceiling(Sum(Case UseBaseUn When 'N' Then T1.OpenQty Else 0 End)/" & Pack & "PackUn) PackUnit " & _
														"from " & oTable & " T0 " & _
														"inner join " & oTable1 & " T1 on T1.DocEntry = T0.DocEntry " & _
														"inner join OITM T2 on T2.ItemCode = T1.ItemCode " & _
														"where T0.DocNum = " & Request("txtOrderNum") & " and T1.ItemCode = N'" & saveHTMLDecode(ItemCode, false) & "' " & _
														"Group By T2." & Pack & "PackUn"
													set rp = conn.execute(sql)
													If ShowReqSum Then %>
														<tr>
															<td bgcolor="#66A4FF" colspan="2"><b><font face="Verdana" size="1"><%=getdelOrderCheckAddByPackLngStr("LtxtReqSum")%>&nbsp;</font></b></td>
															<td bgcolor="#82B4FF" class="style1"><font face="Verdana" size="1"><%=FormatNumber(CDbl(rp("Unit")), myApp.QtyDec)%></font></td>
															<td bgcolor="#82B4FF" class="style1"><font face="Verdana" size="1"><%=FormatNumber(CDbl(rp(NumIn & "Unit")), myApp.QtyDec)%></font></td>
															<td bgcolor="#82B4FF" class="style1"><font face="Verdana" size="1"><%=FormatNumber(CDbl(rp("PackUnit")), myApp.QtyDec)%></font></td>
														</tr>
													</table>
													<% End If %></td>
												</tr>
											</table>
										</td>
									</tr>
									<tr>
										<td colspan="2">
										&nbsp;</td>
									</tr>
								</table>
								<% End If %>
							</td>
						</tr>
						<tr>
							<td align="center">
							<input name="btnBackToRep" type="button" value="<%=getdelOrderCheckAddByPackLngStr("LtxtBackToRep")%>" onclick="window.location.href='operaciones.asp?cmd=invChkInOutCheck&txtOrderNum=<%=Request("txtOrderNum")%>';"></td>
						</tr>
					</table>
					</td>
				</tr>
				<input type="hidden" name="txtOrderNum" value='<%=Request("txtOrderNum")%>'>
			</form>
			</table>
			</td>
		</tr>
	</table>
	</center></div>
<script type="text/javascript">
function onScan(ev){
var scan = ev.data;
	document.frmAddByPack.txtItem.value = scan.value;
	document.frmAddByPack.submit();
}
function onSwipe(ev){
}
try
{
document.addEventListener("BarcodeScanned", onScan, false);
document.addEventListener("MagCardSwiped", onSwipe, false);
}
catch(err) {}
function valFrm()
{
	if (!MyIsNumeric(document.frmAddByPack.txtQty.value))
	{
		alert('<%=getdelOrderCheckAddByPackLngStr("DtxtValNumVal")%>');
		document.frmAddByPack.txtQty.focus();
		return false;
	}

	if (document.frmAddByPack.txtItem.value == '')
	{
		alert('<%=getdelOrderCheckAddByPackLngStr("DtxtValEnterValue")%>');
		document.frmAddByPack.txtItem.focus();
		return false;
	}
	return true;
}
<% If Request("ErrItm") <> "" Then %>
alert('<%=getdelOrderCheckAddByPackLngStr("LtxtErrItmCodBar")%>'.replace('{0}', '<%=Replace(Request("ErrItm"), "'", "\'")%>'));
<% End If %>
</script>
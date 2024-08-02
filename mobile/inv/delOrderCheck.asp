<head>
<style type="text/css">
.style1 {
	background-color: #FFFF00;
}
.style2 {
	background-color: #008000;
	color: #FFFF00;
}
.style3 {
	text-align: center;
}
.style4 {
	border-collapse: collapse;
}
</style>
</head>

<% addLngPathStr = "inv/" %>
<!--#include file="lang/delOrderCheck.asp" -->
<% 

If Request("CreateStatus") = "Y" Then CreateIOData

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
 %>
<script type="text/javascript">
<!--
function valFrm()
{
	if (document.frmSearchItm.txtItem.value == '')
	{
		alert('<%=getdelOrderCheckLngStr("LtxtValSearchItm")%>');
		document.frmSearchItm.txtItem.focus();
		return false;
	}
}
function goItem(item)
{
	<% If Request("ShowDataType") = "" Then %>
	document.frmSearchItm.txtItem.value = item;
	document.frmSearchItm.submit();
	<% ElseIf Request("ShowDataType") = "S" Then %>
	window.location.href='operaciones.asp?cmd=invChkInOutCheckSerial&txtOrderNum=<%=Request("txtOrderNum")%>&ItemCode=' + item;
	<% End If %>
}
function onScan(ev){
var scan = ev.data;
	document.frmSearchItm.txtItem.value = scan.value;
	document.frmSearchItm.submit();
}
function onSwipe(ev){
}

try
{
document.addEventListener("BarcodeScanned", onScan, false);
document.addEventListener("MagCardSwiped", onSwipe, false);
}
catch(err) {}

//-->
</script>
<div align="center">
	<center>
	<table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
		<tr>
			<td bgcolor="#9BC4FF">
			<table border="0" cellpadding="0" bordercolor="#111111" width="100%" id="AutoNumber1">
			<form method="post" action="operaciones.asp" name="frmSearchItm" onsubmit="javascript:return valFrm();">
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
					varText = getdelOrderCheckLngStr("LtxtDocCheck")
					varText = Replace(varText, "{0}", rs("ObjDesc"))
					varText = Replace(varText, "{1}", Request("txtOrderNum"))
					Response.WRite varText %>
					</font></b></p>
					</td>
				</tr>
				<tr>
					<td width="100%" align="center">
					<table cellpadding="0" cellspacing="1" border="0" width="100%">
						<tr>
							<td bgcolor="#66A4FF"><font size="1" face="Verdana"><b><% If rs("DocType") = "Sale" Then %><%=getdelOrderCheckLngStr("DtxtClient")%><% ElseIf rs("DocType") = "Buy" Then %><%=getdelOrderCheckLngStr("DtxtProvider")%><% End If %></b></font></td>
							<td bgcolor="#82B4FF"><font size="1" face="Verdana"><%=rs("CardCode")%></font></td>
						</tr>
						<tr>
							<td bgcolor="#66A4FF"><font size="1" face="Verdana"><b><%=getdelOrderCheckLngStr("DtxtName")%></b></font></td>
							<td bgcolor="#82B4FF"><font size="1" face="Verdana"><%=rs("CardName")%></font></td>
						</tr>
						<tr>
							<td bgcolor="#66A4FF"><font size="1" face="Verdana"><b><%=getdelOrderCheckLngStr("LtxtStatus")%> (<%=getdelOrderCheckLngStr("DtxtItems")%>)</b></font></td>
							<td bgcolor="#82B4FF" class="style3"><b><font face="Verdana" size="1">&nbsp;<% Select Case rs("Status")
						Case "O" %><%=getdelOrderCheckLngStr("DtxtPend")%>
						<% Case "P" %><font color="green"><span class="style1">&nbsp;&nbsp;<%=getdelOrderCheckLngStr("LtxtPartial")%>&nbsp;&nbsp;</span></font>
						<% Case "C" %><span class="style2">&nbsp;&nbsp;<%=getdelOrderCheckLngStr("LtxtCompleted")%>&nbsp;&nbsp;</span>
						<% Case "E" %><font color="#CC3300"><%=getdelOrderCheckLngStr("LtxtOverload")%></font>
						<% End Select %></font></b></td>
						</tr>
						<% If rs("HasSerial") = "Y" Then %>
						<tr>
							<td bgcolor="#66A4FF"><font size="1" face="Verdana"><b><%=getdelOrderCheckLngStr("LtxtStatus")%> (<%=getdelOrderCheckLngStr("LtxtSerialNum")%>)</b></font></td>
							<td bgcolor="#82B4FF" class="style3"><b><font face="Verdana" size="1">&nbsp;<% Select Case rs("SerialStatus")
						Case "O" %><%=getdelOrderCheckLngStr("DtxtPend")%>
						<% Case "P" %><font color="green"><span class="style1">&nbsp;&nbsp;<%=getdelOrderCheckLngStr("LtxtPartial")%>&nbsp;&nbsp;</span></font>
						<% Case "C" %><span class="style2">&nbsp;&nbsp;<%=getdelOrderCheckLngStr("LtxtCompleted")%>&nbsp;&nbsp;</span>
						<% End Select %></font></b></td>
						</tr>
						<% End If %>
						<tr>
							<td bgcolor="#66A4FF"><font size="1" face="Verdana"><b><%=getdelOrderCheckLngStr("LtxtProgress")%></b></font></td>
							<td bgcolor="#82B4FF"><font size="1" face="Verdana"><%=FormatNumber(CDbl(rs("Progress")), myApp.PercentDec)%>%</font></td>
						</tr>
						<tr>
							<td bgcolor="#66A4FF"><font size="1" face="Verdana"><b><%=getdelOrderCheckLngStr("LtxtTotalProgVolume")%></b></font></td>
							<td bgcolor="#82B4FF"><font size="1" face="Verdana"><%=rs("TotalVolProg")%></font></td>
						</tr>
						<tr>
							<td bgcolor="#66A4FF"><font size="1" face="Verdana"><b><%=getdelOrderCheckLngStr("LtxtTotalVolume")%></b></font></td>
							<td bgcolor="#82B4FF"><font size="1" face="Verdana"><%=rs("TotalVol")%></font></td>
						</tr>
					</table>
					</td>
				</tr>
				<tr>
					<td width="100%" height="5px"></td>
				</tr>
				<tr>
					<td width="100%">
					<p align="center"><font size="1" face="Verdana"><%=getdelOrderCheckLngStr("LtxtItemSearchNote")%></font></p>
					</td>
				</tr>
				<tr>
					<td width="100%">
					<table border="0" cellpadding="0" cellspacing="1" bordercolor="#111111" width="100%" id="AutoNumber2">
						<tr>
							<td width="100%">
							<div align="center">
								<center>
								<table border="0" cellpadding="0" cellspacing="1" style="width: 100%;" bordercolor="#111111" id="AutoNumber3" class="style4">
									<tr>
										<td width="100%" class="style3">
										<input type="text" name="txtItem" size="20"></td>
									</tr>
									<tr>
										<td width="100%" height="5px"></td>
									</tr>
									<tr>
										<td width="100%">
										<p align="center">
										<input type="submit" name="btnSearch" value="<%=getdelOrderCheckLngStr("DbtnSearch")%>"></p>
										</td>
									</tr>
									<tr>
										<td width="100%" height="5px"></td>
									</tr>
									<tr>
										<td width="100%">
										<p align="center">
										<input type="button" name="btnAddByPack" value="<%=getdelOrderCheckLngStr("LtxtAddByPack")%>" onclick="window.location.href='?cmd=invChkInOutAddByPack&txtOrderNum=<%=Request("txtOrderNum")%>';"></p>
										</td>
									</tr>
									<tr>
										<td width="100%" height="5px"></td>
									</tr>
									<tr>
										<td width="100%">
										<p align="center"><%  If Request("ShowData") <> "Y" Then 
											btnDesc = getdelOrderCheckLngStr("LtxtViewDataReport")
										Else
											btnDesc = getdelOrderCheckLngStr("LtxtHideDataReport")
										End If %><input type="button" name="btnRep" value="<%=btnDesc%>" onclick="javascript:window.location.href='?cmd=invChkInOutCheck&txtOrderNum=<%=Request("txtOrderNum")%>&ShowData=<% If Request("ShowData") <> "Y" Then Response.Write "Y" %>'"></p>
										</td>
									</tr>
									<tr>
										<td width="100%" height="5px"></td>
									</tr>
								</table>
								</center></div>
							</td>
						</tr>
						<% If Request("ShowData") = "Y" Then %>
						<% If HasSerial Then %>
						<tr>
							<td>
							<table cellpadding="0" cellspacing="0" border="0">
								<tr>
									<td><input type="radio" name="ShowDataType" <% If Request("ShowDataType") = "" Then %>checked disabled<% End If %> id="ShowDataTypeItem" value="" onclick="javascript:window.location.href='?cmd=invChkInOutCheck&txtOrderNum=<%=Request("txtOrderNum")%>&ShowData=Y&Status=<%=Request("Status")%>'"></td>
									<td><label for="ShowDataTypeItem"><font size="1" face="Verdana"><%=getdelOrderCheckLngStr("DtxtItems")%></font></label></td>
									<td><input type="radio" name="ShowDataType" <% If Request("ShowDataType") = "S" Then %>checked disabled<% End If %> id="ShowDataTypeSerial" value="S" onclick="javascript:window.location.href='?cmd=invChkInOutCheck&txtOrderNum=<%=Request("txtOrderNum")%>&ShowData=Y&ShowDataType=S&Status=<%=Request("Status")%>'"></td>
									<td><label for="ShowDataTypeSerial"><font size="1" face="Verdana"><%=getdelOrderCheckLngStr("LtxtSerialNum")%></font></label></td>
								</tr>
							</table>
               				</td>
               			</tr>
               			<% End IF %>
						<tr>
							<td>
							<select size="1" name="cmbStatus" style="font-size:10px; font-family:Verdana" onchange="javascript:window.location.href='?cmd=invChkInOutCheck&txtOrderNum=<%=Request("txtOrderNum")%>&ShowData=Y&ShowDataType=<%=Request("ShowDataType")%>&Status=' + this.value">
							<option value=""><%=getdelOrderCheckLngStr("DtxtPend")%>/<%=getdelOrderCheckLngStr("LtxtPartial")%></option>
							<option <% If Request("Status") = "O" or Request("Status") = "E" and Request("ShowDataType") = "S" Then %>selected<% End If %> value="O"><%=getdelOrderCheckLngStr("DtxtPend")%></option>
							<option <% If Request("Status") = "P" Then %>selected<% End If %> value="P"><%=getdelOrderCheckLngStr("LtxtPartial")%></option>
							<option <% If Request("Status") = "C" Then %>selected<% End If %> value="C"><%=getdelOrderCheckLngStr("LtxtCompleted")%></option>
							<% If Request("ShowDataType") <> "S" Then %><option <% If Request("Status") = "E" Then %>selected<% End If %> value="E"><%=getdelOrderCheckLngStr("LtxtOverload")%></option><% End If %>
							<option <% If Request("Status") = "A" Then %>selected<% End If %> value="A"><%=getdelOrderCheckLngStr("DtxtAll")%></option>
               				</select>
               				</td>
               			</tr>
               			<tr>
               				<td>
							<table style="width: 100%">
								<tr>
									<td bgcolor="#66A4FF" class="style3"><font size="1" face="Verdana"><b><%=getdelOrderCheckLngStr("DtxtItem")%></b></font></td>
									<td bgcolor="#66A4FF" class="style3"><font size="1" face="Verdana"><b><%=getdelOrderCheckLngStr("LtxtStatus")%></b></font></td>
									<td bgcolor="#66A4FF" class="style3"><font size="1" face="Verdana"><b><%=getdelOrderCheckLngStr("LtxtVol")%></b></font></td>
								</tr>
								<% 
								set cmd = Server.CreateObject("ADODB.Command")
								cmd.ActiveConnection = connCommon
								cmd.CommandType = adCmdStoredProc
								cmd.CommandText = "DBOLKGetIOCheckLinesData" & Session("ID")
								cmd.Parameters.Refresh
								cmd("@ObjectCode") = Session("ObjCode")
								cmd("@Type") = Session("Type")
								cmd("@DocNum") = Request("txtOrderNum")
								cmd("@LogNum") = Session("IORetVal")
								cmd("@WhsCode") = Session("bodega")
								If Request("ShowDataType") <> "" Then cmd("@DataType") = Request("ShowDataType")
								rs.close
								rs.open cmd, , 3, 1
								If Request("Status") <> "" and Request("Status") <> "A" Then 
									rs.Filter = "Status = '" & Request("Status") & "'"
								ElseIf Request("Status") = "" Then
									rs.Filter = "Status = 'O' or Status = 'P'"
								End If
								If Not rs.Eof Then
								do while not rs.eof %>
								<tr>
									<td bgcolor="#82B4FF"><a href="javascript:goItem('<%=rs("ItemCode")%>');"><font size="1" face="Verdana" color="#000000"><%=rs("ItemCode")%></font></a></td>
									<td bgcolor="#82B4FF" align="center"><b><font face="Verdana" size="1"><% Select Case rs("Status")
									Case "O" %><%=getdelOrderCheckLngStr("DtxtPend")%>
									<% Case "P" %><font color="green"><span class="style1">&nbsp;&nbsp;<%=getdelOrderCheckLngStr("LtxtPartial")%>&nbsp;&nbsp;</span></font>
									<% Case "C" %><span class="style2">&nbsp;&nbsp;<%=getdelOrderCheckLngStr("LtxtCompleted")%>&nbsp;&nbsp;</span>
									<% Case "E" %><font color="#CC3300"><%=getdelOrderCheckLngStr("LtxtOverload")%></font>
									<% End Select %></font></b></td>
									<td bgcolor="#82B4FF" align="right"><font size="1" face="Verdana"><%=rs("Volume")%>&nbsp;</font></td>
								</tr>
								<% rs.movenext
								loop
								Else %>
								<tr>
									<td bgcolor="#82B4FF" colspan="3" align="center"><font size="1" face="Verdana" color="#000000"><%=getdelOrderCheckLngStr("DtxtNoData")%></font></td>
								</tr>
								<% End If %>
							</table>
							</td>
						</tr>
						<tr>
							<td width="100%" height="5px">
							</td>
						</tr>
						<% End If
						If ObjectCode <> -1 Then %>
						<tr>
							<td width="100%">
							<p align="center">
							<input type="button" name="btnConfirm" value="<%=getdelOrderCheckLngStr("LtxtEndCheckProcess")%>" onclick="javascript:doConfirm();"></p>
							</td>
						</tr>
						<% End If %>
					</table>
					</td>
				</tr>
				<input type="hidden" name="cmd" value="searchInvChkInOutCheckItem">
				<input type="hidden" name="txtOrderNum" value='<%=Request("txtOrderNum")%>'>
			</form>
			</table>
			</td>
		</tr>
	</table>
	</center></div>
<script language="javascript">
function doConfirm()
{
	var status = '<%=Status%>';
	switch (status)
	{
		case 'O':
			alert('<%=getdelOrderCheckLngStr("LtxtAlertOpenDoc")%>');
			return;
			break;
		case 'P':
			<% If Not EnablePartial Then %>
			alert('<%=getdelOrderCheckLngStr("LtxtAlertPartialDoc")%>');
			return;
			<% End If %>
			break;
		case 'E':
			<% If Not AllowOverload Then %>
			alert('<%=getdelOrderCheckLngStr("LtxtAlertOverloadDoc")%>');
			return;
			<% End If %>
			break;
	}
	<% If HasSerial and ChkSerial <> "N" Then %>
	var serialStatus = '<%=SerialStatus%>';
	if (serialStatus == 'O' || serialStatus == 'P')
	{
		<% If ChkSerial = "E" Then %>
		alert('<%=getdelOrderCheckLngStr("LtxtIncSerialErr")%>');
		return;
		<% ElseIf ChkSerial = "C" Then %>
		if (!confirm('<%=getdelOrderCheckLngStr("LtxtIncSerialConf")%>')) return;
		<% End If %>
	}
	<% End If %>
	<% If ObjectCode = 17 and Status = "P" Then %>
	if (!confirm('<%=getdelOrderCheckLngStr("LtxtConfUpdateOrder")%>')) return;
	<% End If %>
	window.location.href='?cmd=invChkInOutCheckSubmit&txtOrderNum=<%=Request("txtOrderNum")%>';
}
</script>
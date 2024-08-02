<!--#include file="clientInc.asp"-->
<!--#include file="agentTop.asp"-->
<% If userType <> "V" Then Response.Redirect "default.asp" %>
<% If Not myApp.EnableOITM Then Response.Redirect "unauthorized.asp" %>
<% addLngPathStr = "" %>
<!--#include file="lang/agentItemSubmit.asp" -->
<% 
If not myAut.GetObjectProperty(4, "C") and Request("Confirm") <> "Y" or Request("isUpdate") = "True" Then goStatus = "C" Else goStatus = "H"
If Session("ItmRetVal") <> "" Then RetVal = Session("ItmRetVal") Else RetVal = Session("ConfItmRetVal")
If Request("doSubmit") = "Y" Then
	set rs = Server.CreateObject("ADODB.RecordSet")
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKGetLogStatus"
	cmd.Parameters.Refresh()
	cmd("@LogNum") = RetVal
	set rs = cmd.execute()
	If rs("Status") = "R" Then
		Session("NotifyAdd") = True
	
		cmd.CommandText = "DBOLKExecuteLog"
		cmd.Parameters.Refresh()
		cmd("@LogNum") = Session("ItmRetVal")
		cmd("@SlpCode") = Session("vendid")
		cmd("@branchIndex") = Session("branch")
		If goStatus = "H" Then cmd("@Confirm") = "Y"
		cmd.execute()

		If goStatus = "H" Then 
			cmd.CommandText = "DBOLKCreateUAFControl" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@UserType") = userType
			cmd("@ExecAt") = "A1"
			cmd("@ObjectEntry") = Session("ItmRetVal")
			cmd("@AgentID") = Session("vendid")
			cmd("@LanID") = Session("LanID")
			cmd("@branch") = Session("branch")
			cmd("@SetLogNumConf") = "Y"
			cmd.Execute()
		End If

		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKPostObjectCreation" & Session("ID")
		cmd.Parameters.Refresh
		cmd("@sessiontype") = "A"
		cmd("@transtype") = "A"
		cmd("@object") = 4 
		cmd("@LogNum") = Session("ItmRetVal")
		cmd("@CurrentSlpCode") = Session("vendid")
		cmd("@Branch") = Session("branch")
		cmd.execute()
	End If
	doItemNextOp()
Else
	ShowNewItem()
End If

Sub doItemNextOp()
	If goStatus <> "H" Then
		doSubmitItemWait()
	Else
		RetVal = Session("ItmRetVal")
		ShowNewItem()
	End If
End Sub

Sub doSubmitItemWait()
	set mySubmit = new SubmitControl
	mySubmit.EnableRunInBackground = True
	mySubmit.LogNum = RetVal 
	mySubmit.LogNumID = "ItmRetVal"
	If Request("isUpdate") = "True" Then
		mySubmit.TransactionOkMessage = getagentItemSubmitLngStr("LtxtConfUpdItm")
	Else
		mySubmit.TransactionOkMessage = getagentItemSubmitLngStr("LtxtConfAddItm")
		mySubmit.EndButtonDescription = getagentItemSubmitLngStr("LtxtCreateNewItem")
		mySubmit.EndButtonFunction = "window.location.href='addItem/goNewItem.asp?AddPath=../';"
		mySubmit.SecondButtonDescription = getagentItemSubmitLngStr("LtxtView") & " " & getagentItemSubmitLngStr("DtxtItem")
		mySubmit.SecondButtonFunction = "viewDetails('{0}');"
		mySubmit.RunInBackgroundRedir = "agentItemSubmit.asp?RetVal=" & Session("ItmRetVal") & "&Confirm=Y&bg=Y"
	End If
	mySubmit.GenerateSubmit %>
<script type="text/javascript">
function Start(page) 
{
	OpenWin = this.open(page, 'objDetails', 'toolbar=no,menubar=no,location=no,scrollbars=yes,resizable=yes, height=600,width=800');
	OpenWin.focus()
}

function viewDetails(itemCode)
{
	Start('');
	doMyLink('addItem/itmConfDetail.asp', 'ItemCode=' + itemCode + '&pop=Y&AddPath=../', 'objDetails');
}
</script>
<% End Sub

Sub ShowNewItem
	
set rs = server.CreateObject("ADODB.RecordSet")
Confirm = goStatus = "H" or Request("bg") = "Y"
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetItmDetails" & Session("ID")
cmd.Parameters.Refresh()
cmd("@LanID") = Session("LanID")
cmd("@LogNum") = RetVal
If Not Confirm Then
	cmd("@ItmType") = 4
Else
	cmd("@ItmType") = -2
End If
set rs = cmd.execute()
EnableSDK = rs("EnableSDK")
 %>
<table border="0" cellpadding="0" width="100%" id="table6">
	<tr class="GeneralTlt">
		<td><% If Not Confirm Then %><%=getagentItemSubmitLngStr("LtxtAddItm")%><% Else %><%=getagentItemSubmitLngStr("LtxtItmConfirm")%><% End If %></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table7">
			<tr>
				<td width="15%" class="GeneralTblBold2"><%=getagentItemSubmitLngStr("DtxtCode")%></td>
				<td width="29%" class="GeneralTbl"><%=rs("ItemCode")%> </td>
				<td width="54%" colspan="2" align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>" class="GeneralTbl">
				<table border="0" cellpadding="0" cellspacing="1" bordercolor="#111111" id="table14">
					<tr>
						<td><font size="1" face="Verdana"><b>
						<input disabled style="background: background-image; border: 0px solid" type="checkbox" <% if rs("prchseitem") = "y" then %>checked<% end if %>></b></font></td>
						<td><b><font size="1" face="Verdana"><%=getagentItemSubmitLngStr("LtxtPurItem")%></font></b></td>
						<td><font size="1" face="Verdana"><b>
						<input disabled style="background: background-image; border: 0px solid" type="checkbox" <% if rs("sellitem") = "y" then %>checked<% end if %>></b></font></td>
						<td><b><font size="1" face="Verdana"><%=getagentItemSubmitLngStr("LtxtSalItem")%></font></b></td>
						<td><font size="1" face="Verdana"><b>
						<input disabled style="background: background-image; border: 0px solid" type="checkbox" <% if rs("invntitem") = "y" then %>checked<% end if %>></b></font></td>
						<td><b><font size="1" face="Verdana"><%=getagentItemSubmitLngStr("LtxtInvItem")%></font></b></td>
					</tr>
				</table>
				</td>
			</tr>
			<tr>
				<td width="15%" class="GeneralTblBold2"><%=getagentItemSubmitLngStr("LtxtDesc1")%></td>
				<td colspan="3" class="GeneralTbl"><%=rs("ItemName")%> </td>
			</tr>
			<tr>
				<td width="15%" class="GeneralTblBold2"><%=getagentItemSubmitLngStr("LtxtDesc2")%></td>
				<td colspan="3" class="GeneralTbl"><%=rs("FrgnName")%> </td>
			</tr>
			<tr>
				<td width="15%" class="GeneralTblBold2"><%=txtAlterGrp%></td>
				<td width="29%" class="GeneralTbl"><%=rs("ItmsGrpNam")%> </td>
				<td width="6%" class="GeneralTblBold2"><%=txtAlterFrm%></td>
				<td width="49%" class="GeneralTbl"><%=rs("FirmName")%> </td>
			</tr>
			<tr>
				<td width="15%" class="GeneralTblBold2"><%=getagentItemSubmitLngStr("LtxtBarCod")%></td>
				<td width="83%" colspan="3" class="GeneralTbl"><%=rs("CodeBars")%> </td>
			</tr>
		</table>
		</td>
	</tr>
	<tr class="GeneralTblBold2">
		<td>
		<p align="center"><%=getagentItemSubmitLngStr("DtxtAddData")%></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table8">
			<tr>
				<td class="GeneralTblBold2"><%=getagentItemSubmitLngStr("LtxtPurUn")%> / <%=getagentItemSubmitLngStr("DtxtQty")%></td>
				<td class="GeneralTbl">&nbsp;<%=rs("BuyUnitMsr")%>&nbsp;<%=rs("NumInBuy")%></td>
				<td class="GeneralTblBold2"><%=getagentItemSubmitLngStr("LtxtPurPackUn")%> / <%=getagentItemSubmitLngStr("DtxtQty")%></td>
				<td class="GeneralTbl">&nbsp;<%=rs("PurPackMsr")%>&nbsp;<%=rs("PurPackUn")%></td>
			</tr>
			<tr>
				<td class="GeneralTblBold2"><%=getagentItemSubmitLngStr("LtxtSalUn")%> / <%=getagentItemSubmitLngStr("DtxtQty")%></td>
				<td class="GeneralTbl">&nbsp;<%=rs("SalUnitMsr")%>&nbsp;<%=rs("NumInSale")%></td>
				<td class="GeneralTblBold2"><%=getagentItemSubmitLngStr("LtxtSalPackUn")%> / <%=getagentItemSubmitLngStr("DtxtQty")%></td>
				<td class="GeneralTbl">&nbsp;<%=rs("SalPackMsr")%>&nbsp;<%=rs("SalPackUn")%></td>
			</tr>
			<tr>
				<td colspan="5" valign="top">
				<table border="0" cellpadding="0" width="100%" id="table9">
					<tr>
						<td valign="top" width="25%">
						<table border="0" cellpadding="0" width="100%" id="table10">
							<tr>
								<td class="GeneralTblBold2">
								<%=getagentItemSubmitLngStr("DtxtImage")%>
								</td>
							</tr>
							<tr>
								<td class="GeneralTbl" height="180">
								<p align="center"><font face="Verdana" size="1">
								<% If IsNull(rs("PicturName")) or Trim(rs("PicturName")) = "" Then Picture = "n_a.gif" Else Picture = rs("PicturName") %>
								<img id="ItemImg" src='pic.aspx?filename=<%=Picture%>&amp;MaxSize=180&amp;dbName=<%=Session("olkdb")%>' border="1" name="ItemImg0"></font></p>
								</td>
							</tr>
						</table>
						</td>
						<td width="75%" valign="top">
						<table border="0" cellpadding="0" width="100%" id="table12">
							<tr class="GeneralTblBold2">
								<td><%=getagentItemSubmitLngStr("DtxtObservations")%></td>
							</tr>
							<tr>
								<td class="GeneralTbl" valign="top" height="180">
								<% If Not IsNull(rs("UserText")) Then %><%=myHTMLEncode(rs("UserText"))%><% End If %></td>
							</tr>
						</table>
						</td>
					</tr>
				</table>
				</td>
			</tr>
		</table>
		</td>
	</tr>
	<% If EnableSDK = "Y" Then
	set rg = Server.CreateObject("ADODB.RecordSet")
	cmd.CommandText = "DBOLKGetUDFGroups" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@LanID") = Session("LanID")
	cmd("@TableID") = "OITM"
	cmd("@UserType") = "V"
	cmd("@OP") = "O"
	set rg = cmd.execute()

	set rc = Server.CreateObject("ADODB.RecordSet")
	cmd.CommandText = "DBOLKGetUDFReadCols" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@LanID") = Session("LanID")
	cmd("@TableID") = "OITM"
	cmd("@UserType") = "V"
	cmd("@OP") = "O"
	rc.open cmd, , 3, 1

		set rd = Server.CreateObject("ADODB.RecordSet")
		
		do while not rg.eof
	 %>
	<tr class="GeneralTblBold2">
		<td>
		<p align="center"><% Select Case CInt(rg("GroupID"))
				Case -1 %><%=getagentItemSubmitLngStr("DtxtUDF")%><%
				Case Else
					Response.Write rg("GroupName")
				End Select %></td>
	</tr>
	<tr>
		<td align="center">
		<table border="0" cellpadding="0" width="100%" id="table13">
			<tr>
				<% 
				arrPos = Split("I,D", ",")
				For i = 0 to 1
				rc.Filter = "GroupID = " & rg("GroupID") & " and Pos = '" & arrPos(i) & "'"
				If not rc.eof then %>
				<td width="50%" valign="top">
				<table border="0" cellpadding="0" width="100%">
					<% do while not rc.eof
			        fldSdk = "U_" & rc("AliasID") %>
					<tr>
						<td width="100" valign="top" class="GeneralTblBold2"><%=rc("Descr")%> </td>
						<td dir="ltr" class="GeneralTbl"><% If rc("TypeID") = "M" and rc("EditType") = "B" Then %><a target="_blank" href="<%=rs(fldSdk)%>"><% End If %>
						<% If rc("TypeID") = "B" Then
				            If Not IsNull(rs(fldSdk)) Then
				            	Select Case rc("EditType")
									Case "R"
										Response.Write FormatNumber(CDbl(rs(fldSdk)),myApp.RateDec)
									Case "S"
										Response.Write FormatNumber(CDbl(rs(fldSdk)),myApp.SumDec)
									Case "P"
										Response.Write FormatNumber(CDbl(rs(fldSdk)),myApp.PriceDec)
									Case "Q"
										Response.Write FormatNumber(CDbl(rs(fldSdk)),myApp.QtyDec)
									Case "%"
										Response.Write FormatNumber(CDbl(rs(fldSdk)),myApp.PercentDec)
									Case "M"
										Response.Write FormatNumber(CDbl(rs(fldSdk)),myApp.MeasureDec)
				            	End Select
				            End If
			            ElseIf rc("TypeID") = "A" and rc("EditType") = "I" Then %>
						<% If IsNull(rs(fldSdk)) or Trim(rs(fldSdk)) = "" Then Picture = "n_a.gif" Else Picture = rs(fldSdk) %>
						<img src='pic.aspx?filename=<%=Picture%>&amp;MaxSize=180&amp;dbName=<%=Session("olkdb")%>' border="0">
						<% Else %> <% If Not IsNull(rs(fldSdk)) Then %><%=Server.HTMLEncode(rs(fldSdk))%><% End If %> <% End If %> <% If rc("TypeID") = "M" and rc("EditType") = "B" Then %></a><% End If %>
						</td>
					</tr>
					<% rc.movenext
			        loop
			        rc.movefirst %>
				</table></td>
				<% End If
				Next %>
			</tr>
		</table>
		</td>
	</tr>
	<% rg.movenext
	loop %>
	<% End If %>
</table>
<% End Sub %>
<!--#include file="agentBottom.asp"-->
<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>

<!--#include file="../Upload/ShadowUploader.asp" -->
<!--#include file="../chkLogin.asp" -->
<!--#include file="lang/cartImport.asp" -->
<html <% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>>

<head>
<% 
set rs = Server.CreateObject("ADODB.recordset")
If userType = "C" Then
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKGetCSSPath" & Session("ID")
	cmd.Parameters.Refresh()
	set rs = cmd.execute()
	SelDes = rs(0)
Else
	SelDes = "0"
End If
%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=getcartImportLngStr("LttlCartImport")%></title>
<!--#include file="../myHTMLEncode.asp"-->
<!--#include file="../lcidReturn.inc" -->
<link rel="stylesheet" type="text/css" href="../design/<%=SelDes%>/style/stylePopUp.css">
<style>
.noborder
{
	border-style:solid; border-width:0; background:background-image;
}
.style1 {
	list-style-type: square;
}
</style>
<script language="javascript">
var txtValTxtFile = '<%=getcartImportLngStr("LtxtValTxtFile")%>';
var txtValNumVal = "<%=getcartImportLngStr("DtxtValNumVal")%>";
var txtValNumMinVal = "<%=getcartImportLngStr("DtxtValNumMinVal")%>";
var txtValNumMaxVal = "<%=getcartImportLngStr("DtxtValNumMaxVal")%>";
var UnEmbPriceSet = <%=JBool(myApp.UnEmbPriceSet)%>;
var txtValMaxQty = '<%=getcartImportLngStr("LtxtValMaxQty")%>';
var txtValSelItms = '<%=getcartImportLngStr("LtxtValSelItms")%>';
</script>
<script language="javascript" src="../generalData.js.asp?dbID=<%=Session("ID")%>&LastUpdate=<%=myApp.LastUpdate%>"></script>
<script language="javascript" src="../general.js"></script>
<script language="javascript" src="cartImport.js"></script>
</head>

<body topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" onload="setTblSet()" onscroll="setTblSet()">

<table border="0" cellpadding="0" cellspacing="1" width="100%" id="table1">
	<tr class="GeneralTblBold2">
		<td><%=getcartImportLngStr("LttlCartImport")%></td>
	</tr>
	<% If Request("doUpload") <> "Y" Then %>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" style="font-family: Verdana; font-size: 10px">
			<form method="POST" enctype="multipart/form-data" action="cartImport.asp?doUpload=Y" onsubmit="javascript:return valFrmUpload(this);" name="frmUploadFile">
			<tr>
				<td class="GeneralTblBold2">
				<%=getcartImportLngStr("LtxtSelDataFile")%> (<%=getcartImportLngStr("LtxtTabDelText")%>)</td>
			</tr>
			<tr>
				<td class="GeneralTbl">
				<input type="file" name="xmlFile" size="58" onchange="javascript:if (this.value.substring(this.value.length-3).toLowerCase() != 'txt'){ alert('<%=getcartImportLngStr("LtxtValTxtFile")%>');this.value=''; };"></td>
			</tr>
			<tr>
				<td class="GeneralTblBold2">
				<input type="submit" value="<%=getcartImportLngStr("LtxtNext")%>" name="btnNext"></td>
			</tr>
			<tr>
				<td>&nbsp;</td>
			</tr>
			<tr>
				<td align="center">
					<table border="0" cellpadding="0" width="90%" style="font-family: Verdana; font-size: 10px; background-color: #FFFFCC; padding: 5px;" align="center">
						<tr>
							<td rowspan="2" style="vertical-align: top; width: 60px; "><img src="../images/guide_info.gif"></td>
							<td><b><u><%=getcartImportLngStr("DtxtGuide")%></u></b></td>
						</tr>
						<tr>
							<td><%
							If userType = "C" Then reqCols = 2 Else reqCols = 4
							Response.Write Replace(getcartImportLngStr("LtxtExplanation"), "{0}", reqCols)%></td>
						</tr>
						<tr>
							<td>&nbsp;</td>
							<td style="padding-right: 10px; padding-left: 10px;">
							<ol>
								<li><%=getcartImportLngStr("DtxtItemCode")%> (<%=getcartImportLngStr("LtxtItemCodeDesc")%>)</li>
								<li><%=getcartImportLngStr("DtxtQty")%></li><% IF userType = "V" Then %>
								<li><%=getcartImportLngStr("DtxtUnit")%> (<%=getcartImportLngStr("LtxtUnitDesc")%>)</li>
									<ul class="style1">
										<li>1: <%=getcartImportLngStr("DtxtBaseUnit")%></li>
										<li>2: <%=getcartImportLngStr("DtxtSalUnit")%></li>
										<li>3: <%=getcartImportLngStr("DtxtPackUnit")%></li>
								</ul>
								<li><%=getcartImportLngStr("DtxtPrice")%> (<%=getcartImportLngStr("LtxtPriceDesc")%>)</li><% End If %>
							</ol>
							</td>
						</tr>
					</table>
				</td>
			</tr>
			</form>
		</table>
		</td>
	</tr>
	<% 
	Else
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKGetObject"
	cmd.Parameters.Refresh()
	cmd("@LogNum") = Session("RetVal")
	set rs = cmd.execute()
	CartObject = rs(0)
	Dim objUpload
	set objUpload = New ShadowUpload
	If objUpload.GetError <> "" Then %>
	<tr class="GeneralTblBold2">
		<td>
		<%="" & getcartImportLngStr("LtxtErrUpdFile") & ": " & objUpload.GetError %>
		</td>
	</tr>
	<% Else
	FileName = objUpload.File(0).FileName
	Call objUpload.File(0).SaveToDisk(Server.MapPath("../temp/"), "cartImport.txt")
	%>
	<form method="POST" name="frmImport" onsubmit="javascript:return valFrmImp();" action="cartImportSubmit.asp">
	<tr>
		<td class="GeneralTblBold2">
		<%=getcartImportLngStr("LtxtSelItems")%></td>
	</tr>
	<tr>
		<td>
			<table border="0" cellpadding="0" cellspacing="1" width="100%" id="table2">
				<tr class="GeneralTblBold2">
					<td width="10">&nbsp;</td>
					<td align="center">
					<%=getcartImportLngStr("DtxtItem")%></td>
					<td align="center">
					<%=getcartImportLngStr("DtxtQty")%></td>
					<td align="center">
					<%=getcartImportLngStr("DtxtUnit")%></td>
					<td align="center">
					<%=getcartImportLngStr("DtxtPrice")%></td>
					<td align="center">
					<%=getcartImportLngStr("DtxtTotal")%></td>
					<td align="center">
					<%=getcartImportLngStr("DtxtValid")%></td>
				</tr>
				<% 
				set connImp = Server.CreateObject("ADODB.Connection")
				set rsImp = Server.CreateObject("ADODB.RecordSet")
				connImp.open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
							"Data Source=" & Server.MapPath("../temp/") & ";" & _
							"Extended Properties=""Text;HDR=NO;"""
				sql = "select * from cartImport.txt"
				rsImp.open sql, connImp
				LineNum = 0
				do while not rsImp.eof
				LineNum = LineNum + 1
				ItemCode = rsImp(0)
				
				If userType = "V" and rsImp.Fields.Count = 4 or userType = "C" and rsImp.Fields.Count = 2 Then
					Select Case userType
						Case "C"
							Quantity = rsImp(1)
							Price = 0
						Case "V"
							Quantity = rsImp(1)
							SaleUnit = rsImp(2)
							Price = rsImp(3)
					End Select
				Else
					ItemCode = ""
				End If
				
				If ItemCode <> "" and IsNumeric(Quantity) and IsNumeric(myApp.GetSaleUnit) and IsNumeric(Price) Then
				
				set cmd = Server.CreateObject("ADODB.Command")
				cmd.ActiveConnection = connCommon
				cmd.CommandType = &H0004
				cmd.CommandText = "DBOLKCheckItemImport" & Session("ID")
				cmd.Parameters.Refresh()
				cmd("@ItemCode") = ItemCode
				cmd("@DiscPrice") = CDbl(getNumericOut(Price))
				cmd("@UserType") = userType
				cmd("@branch") = Session("branch")
				cmd("@SlpCode") = Session("vendid")
				cmd("@CardCode") = Session("UserName")
				cmd("@PriceList") = Session("PriceList")
				cmd("@Quantity") = CDbl(getNumericOut(Quantity))
				cmd("@LogNum") = Session("RetVal")
				set rs = cmd.execute()

				addItems = False
				If Not rs.Eof Then
				addItems = True
				IsValid = rs("Verfy") = "Y" or CartObject = 23
				If Not IsValid Then
					ErrMsg = Replace("" & getcartImportLngStr("LtxtValQty") & "", "{0}", rs("VerfyMaxQty"))
				Else
				End If %>
				<tr class="GeneralTbl">
					<td width="10">
					<p align="left">
					<input type="hidden" name="ItemCode<%=LineNum%>" value="<%=ItemCode%>">
					<input type="checkbox" checked name="LineNum" id="LineNum<%=LineNum%>" value="<%=LineNum%>" class="noborder" onclick="javascript:chkCheckAll();"></td>
					<td><label for="LineNum<%=LineNum%>"><%=ItemCode%></label>
					</td>
					<td>
					<input type="text" name="Quantity<%=LineNum%>" value="<%=FormatNumber(Quantity, myApp.QtyDec)%>" style="text-align: right; width: 100%;" onkeydown="return valKeyNumDec(event);" onfocus="javascript:this.select();" onkeydown="return valKeyNumDec(event);" onchange="javascript:ChkNum(this, '<%=FormatNumber(0.000001, myApp.QtyDec)%>', '<%=FormatNumber(CDbl(rs("VerfyMaxQty")), myApp.QtyDec)%>', '<%=FormatNumber(Quantity, myApp.QtyDec)%>', <%=myApp.QtyDec%>);setTotal('<%=LineNum%>');">
					<input type="hidden" name="MaxQty<%=LineNum%>" value="<%=FormatNumber(CDbl(rs("VerfyMaxQty")), myApp.QtyDec)%>">
					</td>
					<td align="center"><span dir="ltr"><% Select Case myApp.GetSaleUnit
						Case 1 %><%=getcartImportLngStr("LtxtUn")%>
					<%	Case 2 %><%=rs("SalUnitMsr")%><% If myApp.GetShowQtyInUn Then %>(<%=rs("NumInSale")%>)<% End If %>
					<%	Case 3 %><%=rs("SalPackMsr")%><% If myApp.GetShowQtyInUn Then %>(<%=rs("SalPackUn")%>)<% End If %>
					<%
					End Select %></span>
					<input type="hidden" name="SaleUnit<%=LineNum%>" value="<%=myApp.GetSaleUnit%>">
					</td>
					<td>
					<p align="center">
					<% Select Case userType
						Case "C" %><%=FormatNumber(CDbl(rs("Price")), myApp.PriceDec)%><input type="hidden" name="Price<%=LineNum%>" value="<%=FormatNumber(CDbl(rs("Price")), myApp.PriceDec)%>">
					<% 	Case "V" %><input type="text" name="Price<%=LineNum%>" value="<%=FormatNumber(CDbl(rs("Price")), myApp.PriceDec)%>" style="text-align: right; width: 100%" onfocus="javascript:this.select();" onchange="javascript:ChkNum(this, '<%=FormatNumber(0.000001, myApp.PriceDec)%>', '', '<%=FormatNumber(CDbl(rs("Price")), myApp.PriceDec)%>', <%=myApp.PriceDec%>);setTotal('<%=LineNum%>');">
					<% End Select %></td>
					<td id="LineTotal<%=LineNum%>" align="right">
					<%
					If SaleUnit = 3 and myApp.UnEmbPriceSet Then PriceBy = CDbl(rs("SalPackUn")) Else PriceBy = 1
					Response.Write FormatNumber(CDbl(rs("Price"))*CDbl(Quantity)*PriceBy, myApp.SumDec) %>
					</td>
					<td align="center">
					<p align="center">
					<% If Not IsValid Then %><a href="#" onclick="javascript:alert('<%=Replace(ErrMsg, "'", "\'")%>');"><% End If %>
					<font color="#<% If IsValid Then %>31659C<% Else %>FF0000<% End If %>" size="3" face="Wingdings"><% If IsValid Then %>&#252;<% Else %>&#251;<% End If %></font>
					<% If Not IsValid Then %></a><% End If %>
					</td>
					<input type="hidden" name="NumInSale<%=LineNum%>" value="<%=rs("NumInSale")%>">
					<input type="hidden" name="SalPackUn<%=LineNum%>" value="<%=rs("SalPackUn")%>">
				</tr>
				<% Else %>
				<tr class="GeneralTbl">
					<td colspan="8" align="center"><font color="#FF0000"><%=Replace(getcartImportLngStr("LtxtItemNotFound"), "{0}", ItemCode)%></font></td>
				</tr>
				<% End If
				Else %>
				<tr class="GeneralTbl">
					<td colspan="8" align="center"><font color="#FF0000"><%=Replace(getcartImportLngStr("LtxtLineStructErr"), "{0}", LineNum)%></font></td>
				</tr>
				<% End If %>
				<% rsImp.movenext
				loop
				If addItems Then %>
				<tr class="GeneralTbl">
					<td colspan="2">
					<p>
					<input type="checkbox" checked name="chkAllItems" value="Y" id="chkAllItems" class="noborder" onclick="javascript:chkAll(this.checked);">
					<label for="chkAllItems"><%=getcartImportLngStr("DtxtAll")%></label></td>
					<td>
					&nbsp;</td>
					<td>
					&nbsp;</td>
					<td>
					&nbsp;</td>
					<td>
					<p>
					&nbsp;</td>
					<td>
					&nbsp;</td>
				</tr>
				<% End If %>
			</table>
		</td>
	</tr>
	<script language="javascript">
	CartObject = <%=CartObject%>;
	</script>
	<% End If %>
	<% End If %>
	<tr>
		<td>
		<table border="0" cellpadding="0" class="GeneralTbl" cellspacing="1" width="100%" id="tblImport" style="position: absolute;  z-index: 1;">
			<tr>
				<% If Request.QueryString("doUpload") = "Y" and addItems Then %><td>
				<input type="submit" value="<%=getcartImportLngStr("DtxtImport")%>" name="btnImport"></td><% End If %>
				<td align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
				<input type="button" value="<%=getcartImportLngStr("DtxtCancel")%>" name="btnClose" onclick="javascript:window.close();"></td>
			</tr>
		</table>
		</td>
	</tr>
	<% If Request.QueryString("doUpload") = "Y" Then %>
	</form>
	<% End If %>
</table>

</body>

</html>
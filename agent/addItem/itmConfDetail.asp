<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="../chkLogin.asp" -->
<!--#include file="../authorizationClass.asp"-->
<!--#include file="lang/itmConfDetail.asp" -->
<html <% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>>

<%
Dim varx
varx = "0"
%>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=getitmConfDetailLngStr("LttlItmDetailsConf")%></title>
<%
Dim myAut
set myAut = New clsAuthorization

Dim DocName

%>
<!--#include file="../loadAlterNames.asp"-->
<link rel="stylesheet" type="text/css" href="../design/0/style/stylenuevo.css">
</head>

<body>

<% 

hasAut = myAut.GetObjectProperty(4, "V")
If hasAut Then

set rd = Server.CreateObject("ADODB.RecordSet")
cmd.CommandText = "DBOLKGetItmDetails" & Session("ID")
cmd.Parameters.Refresh()
cmd("@LanID") = Session("LanID")
If Request("ItemCode") <> "" Then
	cmd("@ItmType") = 4
	cmd("@ItemCode") = Request("ItemCode")
Else
	cmd("@ItmType") = -2
	cmd("@LogNum") = Request("DocEntry")
End If
set rs = cmd.execute()
EnableSDK = rs("EnableSDK")

%>
<table border="0" cellpadding="0" width="100%" id="table6">
	<tr class="GeneralTlt">
		<td><% If Request("ItemCode") <> "" Then %><%=getitmConfDetailLngStr("LtxtItemDet")%><% Else %><%=getitmConfDetailLngStr("LttlItmDetailsConf")%> #<%=Request("DocEntry")%><% End If %></td>
	</tr>
	<tr>
		<td height="89">
		<table border="0" cellpadding="0" width="100%" id="table7">
			<tr>
				<td width="15%" class="GeneralTblBold2"><%=getitmConfDetailLngStr("DtxtCode")%></td>
				<td width="29%" class="GeneralTbl"><%=rs("ItemCode")%> </td>
				<td width="54%" colspan="2" align="right" class="GeneralTbl">
				<table border="0" cellpadding="0" cellspacing="1" bordercolor="#111111" id="table14">
					<tr>
						<td><font size="1" face="Verdana"><b>
						<input disabled style="background: background-image; border: 0px solid" type="checkbox" <% if rs("prchseitem") = "Y" then %>checked<% end if %>></b></font></td>
						<td><b><font size="1" face="Verdana"><%=getitmConfDetailLngStr("LtxtPurItem")%></font></b></td>
						<td><font size="1" face="Verdana"><b>
						<input disabled style="background: background-image; border: 0px solid" type="checkbox" <% if rs("sellitem") = "Y" then %>checked<% end if %>></b></font></td>
						<td><b><font size="1" face="Verdana"><%=getitmConfDetailLngStr("LtxtSalItem")%></font></b></td>
						<td><font size="1" face="Verdana"><b>
						<input disabled style="background: background-image; border: 0px solid" type="checkbox" <% if rs("invntitem") = "Y" then %>checked<% end if %>></b></font></td>
						<td><b><font size="1" face="Verdana"><%=getitmConfDetailLngStr("LtxtInvItem")%></font></b></td>
					</tr>
				</table>
				</td>
			</tr>
			<tr>
				<td width="15%" class="GeneralTblBold2"><%=getitmConfDetailLngStr("LtxtDesc1")%></td>
				<td colspan="3" class="GeneralTbl"><%=rs("ItemName")%> </td>
			</tr>
			<tr>
				<td width="15%" class="GeneralTblBold2"><%=getitmConfDetailLngStr("LtxtDesc2")%></td>
				<td colspan="3" class="GeneralTbl"><%=rs("FrgnName")%> </td>
			</tr>
			<tr>
				<td width="15%" class="GeneralTblBold2"><%=txtAlterGrp%></td>
				<td width="29%" class="GeneralTbl"><%=rs("ItmsGrpNam")%> </td>
				<td width="6%" class="GeneralTblBold2"><%=txtAlterFrm%></td>
				<td width="49%" class="GeneralTbl"><%=rs("FirmName")%> </td>
			</tr>
			<tr>
				<td width="15%" class="GeneralTblBold2"><%=getitmConfDetailLngStr("LtxtBarCod")%></td>
				<td width="83%" colspan="3" class="GeneralTbl"><%=rs("CodeBars")%> </td>
			</tr>
		</table>
		</td>
	</tr>
	<tr class="GeneralTblBold2">
		<td>
		<p align="center"><%=getitmConfDetailLngStr("DtxtAddData")%></p>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table8">
			<tr>
				<td class="GeneralTblBold2"><%=getitmConfDetailLngStr("LtxtPurUn")%> / <%=getitmConfDetailLngStr("DtxtQty")%></td>
				<td class="GeneralTbl">&nbsp;<%=rs("BuyUnitMsr")%>&nbsp;<%=rs("NumInBuy")%></td>
				<td class="GeneralTblBold2"><%=getitmConfDetailLngStr("LtxtPurPackUn")%> / <%=getitmConfDetailLngStr("DtxtQty")%></td>
				<td class="GeneralTbl">&nbsp;<%=rs("PurPackMsr")%>&nbsp;<%=rs("PurPackUn")%></td>
			</tr>
			<tr>
				<td class="GeneralTblBold2"><%=getitmConfDetailLngStr("LtxtSalUn")%> / <%=getitmConfDetailLngStr("DtxtQty")%></td>
				<td class="GeneralTbl">&nbsp;<%=rs("SalUnitMsr")%>&nbsp;<%=rs("NumInSale")%></td>
				<td class="GeneralTblBold2"><%=getitmConfDetailLngStr("LtxtSalPackUn")%> / <%=getitmConfDetailLngStr("DtxtQty")%></td>
				<td class="GeneralTbl">&nbsp;<%=rs("SalPackMsr")%>&nbsp;<%=rs("SalPackUn")%></td>
			</tr>
			<tr>
				<td colspan="5" valign="top">
				<table border="0" cellpadding="0" width="100%" id="table9">
					<tr>
						<td valign="top" width="25%">
						<table border="0" cellpadding="0" width="100%" id="table10">
							<tr>
								<td class="GeneralTblBold2"><%=getitmConfDetailLngStr("DtxtImage")%></td>
							</tr>
							<tr>
								<td class="GeneralTbl" height="180">
								<p><font face="Verdana" size="1">
								<% If IsNull(rs("PicturName")) or Trim(rs("PicturName")) = "" Then Picture = "n_a.gif" Else Picture = rs("PicturName") %>
								<img id="ItemImg0" src="../pic.aspx?filename=<%=Picture%>&MaxSize=180&dbName=<%=Session("olkdb")%>" border="1" name="ItemImg0"></font></p>
								</td>
							</tr>
						</table>
						</td>
						<td width="75%" valign="top">
						<table border="0" cellpadding="0" width="100%" id="table12">
							<tr class="GeneralTblBold2">
								<td><%=getitmConfDetailLngStr("DtxtObservations")%>:</td>
							</tr>
							<tr>
								<td class="GeneralTbl" valign="top" height="180">
								<%=rs("UserText")%></td>
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
		
	do while not rg.eof
	 %>
	<tr class="GeneralTblBold2">
		<td>
		<p align="center"><% Select Case CInt(rg("GroupID"))
				Case -1 %><%=getitmConfDetailLngStr("DtxtUDF")%><%
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
						<img src="../pic.aspx?filename=<%=Picture%>&MaxSize=180&dbName=<%=Session("olkdb")%>" border="0">
						<% Else %> <% If Not IsNull(rs(fldSdk)) Then %><%=rs(fldSdk)%><% End If %> <% End If %> <% If rc("TypeID") = "M" and rc("EditType") = "B" Then %></a><% End If %>
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
	<% Else %>
	<script type="text/javascript">
	alert('<%=getitmConfDetailLngStr("DtxtNoAccessObj")%>'.replace('{0}', '<%=getitmConfDetailLngStr("DtxtItem")%>'));
	window.close();
	</script>
	<% End If %>


</body>
<% set rs = nothing
set rd = nothing
conn.close %></html>
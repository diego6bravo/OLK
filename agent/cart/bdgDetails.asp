<%@ Language=VBScript %>
<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="../chkLogin.asp" -->
<!--#include file="../clearItem.asp"-->
<!--#include file="lang/bdgDetails.asp" -->
<!--#include file="../myHTMLEncode.asp"-->
<!--#include file="../authorizationClass.asp"-->
<html <% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=getbdgDetailsLngStr("LttlInvWhsRep")%></title>
<script language="javascript">
function Start(page, w, h, s) {
OpenWin = this.open(page, "ImageThumb", "toolbar=no,menubar=no,location=no,scrollbars="+s+",resizable=yes, width="+w+",height="+h);
}
</script>
<link rel="stylesheet" type="text/css" href="../design/0/style/stylePopUp.css">
</head>
<%
Dim myAut
set myAut = New clsAuthorization

varx = 0
set rw = Server.CreateObject("ADODB.recordset")
sql = "select OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'ItemName', ItemCode, ItemName) ItemName, PicturName from oitm where itemcode = N'" & saveHTMLDecode(Request("Item"), False) & "'"
set rw = conn.execute(sql)
ItemName = rw("ItemName")
If rw("PicturName") <> "" Then Pic = rw("PicturName") Else Pic = "n_a.gif"
rw.close

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetWarehouses" & Session("ID")
cmd.Parameters.Refresh()
cmd("@LanID") = Session("LanID")
If Not myAut.HasAuthorization(99) Then cmd("@Filter") = Request("WhsCode")
rw.open cmd, , 3, 1
set rs = server.createobject("ADODB.Recordset")

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKSalesItemDetails" & Session("ID")
cmd.Parameters.Refresh()
cmd("@LanID") = Session("LanID")
cmd("@ItemCode") = Request("item")
set rs = cmd.execute()
%>

<body topmargin="0" leftmargin="0">
<table border="0" cellpadding="0" width="579">
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" cellspacing="1">
			<tr class="GeneralTlt">
				<td width="157">&nbsp;<%=getbdgDetailsLngStr("DtxtCode")%>:</td>
				<td><%=Request("Item")%>: <%=ItemName%></td>
			</tr>
			<tr>
				<td width="157" valign="middle">
				<p align="center"><a href="javascript:Start('../thumb/?item=<%=Server.URLEncode(Request("ItemCode"))%>&pop=Y&AddPath=../',529,510,'yes')"><img border="0" src="../pic.aspx?filename=<%=Pic%>&dbName=<%=Session("olkdb")%>"></a></td>
				<td valign="top">
				<table border="0" cellpadding="0" width="100%">
					<tr class="GeneralTblBold2">
						<td colspan="2">
						<p align="center"><%=getbdgDetailsLngStr("LtxtOnHand")%></td>
						<td width="199" colspan="2">
						<p align="center"><%=getbdgDetailsLngStr("LtxtAvl")%></td>
					</tr>
					<tr class="GeneralTbl">
						<td>&nbsp;</td>
						<td>
						<%=getbdgDetailsLngStr("DtxtSAP")%></td>
						<td width="100"><%=getbdgDetailsLngStr("DtxtSAP")%></td>
						<td width="99"><%=getbdgDetailsLngStr("DtxtOLK")%></td>
					</tr>
					<tr class="GeneralTbl">
						<td><%=getbdgDetailsLngStr("DtxtUnit")%></td>
						<td><%=FormatNumber(RS("OnHand"),0)%>&nbsp;</td>
						<td width="100"><%=FormatNumber(RS("DispSAP"),0)%>&nbsp;</td>
						<td width="99"><%=FormatNumber(RS("InvOLKDisp"),0)%>&nbsp;</td>
					</tr>
					<tr class="GeneralTbl">
						<td><%=RS("SalUnitMsr")%>&nbsp;(<%=RS("NumInSale")%>)</td>
						<td><%=FormatNumber(RS("OnHandUnVentSAP"),0)%> (<%=FormatNumber(RS("OnHandSueltoUnVentaSAP"),0)%>)</td>
						<td width="100"><%=FormatNumber(RS("DispUnVentSAP"),0)%> 
						(<%=FormatNumber(RS("DispSueltoUnVentSAP"),0)%>)</td>
						<td width="99"><%=FormatNumber(RS("InvOLKUnVentDisp"),0)%> 
						(<%=FormatNumber(RS("InvOLKSueltoUnVentDisp"),0)%>)&nbsp;</td>
					</tr>
					<tr class="GeneralTbl">
						<td><%=RS("SalPackMsr")%>&nbsp;(<%=RS("SalPackUn")%>)</td>
						<td><%=FormatNumber(RS("OnHandUnEmbSAP"),0)%> (<%=FormatNumber(RS("OnHandSueltoUnEmbSAP"),0)%>)</td>
						<td width="100"><%=FormatNumber(RS("DispUnEmbSAP"),0)%> 
						(<%=FormatNumber(RS("DispSueltoUnEmbSAP"),0)%>)</td>
						<td width="99"><%=FormatNumber(RS("InvOLKUnEmbDisp"),0)%> 
						(<%=FormatNumber(RS("InvOLKSueltoUnEmbDisp"),0)%>)</td>
					</tr>
				</table>
				</td>
			</tr>
		</table>
		</td>
	</tr>
	<tr class="GeneralTbl">
		<td><hr size="1"></td>
	</tr>
	<tr class="GeneralTbl">
		<td>
			<table border="0" cellpadding="0" width="100%">
				<tr>
			  <% while not rw.eof
			  varx = varx + 1
			  	cmd("@WhsCode") = rw("WhsCode")
				  set rs = cmd.execute() %>
					<td>
					<div align="center">
						<table border="0" cellpadding="0" width="272" cellspacing="0">
							<tr class="GeneralTlt">
								<td>
								<table border="0" cellpadding="0" width="272" cellspacing="1">
									<tr class="GeneralTbl">
										<td>
										<table border="0" cellpadding="0" width="100%" cellspacing="1">
											<tr class="GeneralTblBold2">
												<td colspan="4"><%=rw("whsname")%>&nbsp;</td>
											</tr>
											<tr class="GeneralTblBold2">
												<td align="center" width="50%" colspan="2">
												<%=getbdgDetailsLngStr("LtxtOnHand")%></td>
												<td align="center" width="50%" colspan="2">
												<%=getbdgDetailsLngStr("LtxtAvl")%></td>
											</tr>
											<tr class="GeneralTblBold2">
												<td width="25%">&nbsp;</td>
												<td width="25%"><%=getbdgDetailsLngStr("LtxtWHS")%></td>
												<td width="25%"><%=getbdgDetailsLngStr("DtxtSAP")%></td>
												<td width="25%"><%=getbdgDetailsLngStr("DtxtOLK")%></td>
											</tr>
											<tr class="GeneralTblBold2">
												<td width="25%"><%=getbdgDetailsLngStr("DtxtUnit")%></td>
												<td width="25%"><%=FormatNumber(RS("InvBDGWhs"),0)%>&nbsp;</td>
												<td width="25%"><%=FormatNumber(RS("InvBDGDisp"),0)%>&nbsp;</td>
												<td width="25%"><%=FormatNumber(RS("InvOLKBDGDisp"),0)%>&nbsp;</td>
											</tr>
											<tr class="GeneralTblBold2">
												<td width="25%"><%=RS("SalUnitMsr")%><% If myApp.GetShowQtyInUn Then %>&nbsp;(<%=RS("NumInSale")%>)<% End If %></td>
												<td width="25%">
												<%=FormatNumber(RS("InvUnVentBDGWhs"),0)%> 
												(<%=FormatNumber(RS("InvSueltoUnVentBDGWhs"),0)%>)&nbsp;</td>
												<td width="25%"><%=FormatNumber(RS("InvBDGUnVentDisp"),0)%> 
												(<%=FormatNumber(RS("InvBDGSueltoUnVentDisp"),0)%>)</td>
												<td width="25%"><%=FormatNumber(RS("InvOLKBDGUnVentDisp"),0)%> 
												(<%=FormatNumber(RS("InvOLKBDGSueltoUnVentDisp"),0)%>)</td>
											</tr>
											<tr class="GeneralTblBold2">
												<td width="25%"><%=RS("SalPackMsr")%><% If myApp.GetShowQtyInUn Then %>&nbsp;(<%=RS("SalPackUn")%>)<% End If %></td>
												<td width="25%">
										        <%=FormatNumber(RS("InvUnEmbBDGWhs"),0)%> 
												(<%=FormatNumber(RS("InvSueltoUnEmbBDGWhs"),0)%>)&nbsp;</td>
												<td width="25%"><%=FormatNumber(RS("InvBDGUnEmbDisp"),0)%> 
												(<%=FormatNumber(RS("InvBDGSueltoUnEmbDisp"),0)%>)</td>
												<td width="25%"><%=FormatNumber(RS("InvOLKBDGUnEmbDisp"),0)%> 
												(<%=FormatNumber(RS("InvOLKBDGSueltoUnEmbDisp"),0)%>)&nbsp;</td>
											</tr>
										</table>
										</td>
									</tr>
								</table>
								</td>
							</tr>
						</table>
					</div>
					</td>
			  <% if varx = 2 then 
			  response.write "</tr><tr>"
			  varx = 0
			  end if
			  rw.movenext
			  wend %>
				</tr>
			</table>
		</td>
	</tr>
	<tr class="GeneralTbl">
		<td>
		<input type="button" value="<%=getbdgDetailsLngStr("DtxtBack")%>" name="B1" style="float: <% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>" onclick="javascript:history.go(-1);"></td>
	</tr>
</table>
  </body>

</html>
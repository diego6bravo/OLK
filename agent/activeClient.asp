<!--#include file="clientInc.asp"-->
<!--#include file="agentTop.asp"-->
<%
If Session("UserName") = "" Then Response.Redirect "unauthorized.asp"
%>
<% addLngPathStr = "" %>
<!--#include file="lang/activeClient.asp" -->
<% If Session("UserName") = "" then 
	response.redirect "searchClient.asp"
Else
	set rd = Server.CreateObject("ADODB.RecordSet")
	sql = "select AliasID from CUFD where TableID = 'OCRD' and TypeID = 'A' and EditType = 'I'"
	set rd = conn.execute(sql)
	If Not rd.Eof Then
		crdCaseImg = ""
		do while not rd.eof
			If crdCaseImg <> "" Then crdCaseImg = crdCaseImg & " or "
			crdCaseImg = crdCaseImg & "U_" & rd("AliasID") & " is not null "
		rd.movenext
		loop
		crdCaseImg = " Case When " & crdCaseImg & " Then 'Y' Else 'N' End "
	Else
		crdCaseImg = " 'N' "
	End If

	sql = "select T0.CardType, T0.SlpCode, T0.ListNum, T0.Picture, Lower(T0.Country) Country, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRY', 'Name', T1.Code, T1.Name) CountryName, " & _
	"" & crdCaseImg & " MoreImg " & _
	"from ocrd T0 " & _
	"left outer join ocry T1 on T1.Code = T0.Country " & _
	"where T0.CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "'"
	set rs = conn.execute(sql)
	If not myAut.HasAuthorization(60) and CInt(Session("VendId")) <> CInt(rs("SlpCode")) Then Response.Redirect "configErr.asp?errCmd=AsignedSLP"
	CardType = rs("CardType")
	ListNum = rs("ListNum")
	If Not IsNull(rs("Picture")) Then crdPic = rs("Picture") Else crdPic = "pcard.gif"
	Country = rs("Country")
	CountryName = rs("CountryName")
End If
activeClient = True
ShowPendOlk = False %>
<script language="javascript">
<!--
<% If Request("DocFlowErr") <> "" Then %>
function goAdd(Confirm, DocConf)
{
	if (<%=Request("obj")%> == 48) { window.location.href='ventas/newCashInv.asp?DocConf=' + DocConf; }
	else if (<%=Request("obj")%> == 24) { window.location.href='payments/newDocGo.asp?DocConf=' + DocConf; }
	else { window.location.href='ventas/newDocGo.asp?obj=<%=Request("obj")%>&DocConf=' + DocConf; }
}
<% End If %>
function Start(page, w, h, s) {
OpenWin = this.open(page, "ImageThumb", "toolbar=no,menubar=no,location=no,scrollbars="+s+",resizable=yes, width="+w+",height="+h);
}

//-->
</script>
<% 
If Session("rtl") = "" Then addRtl = "" Else addRtl = "rtl/"

Dim myActiveMnu()
mnuCount = -1

If myAut.HasAuthorization(66) Then lnk = "addCard/goEditCard.asp?AddPath=../" Else lnk = "javascript:doMyLink('addCard/crdConfDetailOpen.asp', 'CardCode=" & Session("UserName") & "', '_blank');"
AddActiveMnuItm "client_data_botom", Replace(getactiveClientLngStr("LtxtClientData"), "{0}", txtClient), lnk, "N", False

If 1 = 2 Then
	AddActiveMnuItm "ofertas_botom", txtOferts, "activeClient.asp?open=oferts", "N", Request("open") = "oferts"

	AddActiveMnuItm "pendientes_botom", "" & getactiveClientLngStr("LtxtPendDocs") & "", "javascript:doMyLink('activeClient.asp', 'cCode=" & Session("UserName") & "&open=openedDocs', '');", "N", Request("open") = "openedDocs" or Request("open") = ""
End If

AddActiveMnuItm "mensajes_botom", getactiveClientLngStr("LtxtMessages"), "javascript:doMyLink('newMessage.asp', 'ClientsTo=" & Replace(myHTMLEncode(Session("UserName")), "'", "\'") & "', '')", "N", False

If myApp.EnableOCLG Then
	AddActiveMnuItm "activity_botom", getactiveClientLngStr("DtxtActivity"), "addActivity/goNewActivity.asp?AddPath=../", "N", False
End If

If myApp.EnableOOPR Then
	AddActiveMnuItm "activity_botom", getactiveClientLngStr("DtxtSO"), "addSO/goNewSO.asp?AddPath=../", "N", False
End If

If myApp.EnableOQUT and CardType <> "S" Then
	AddActiveMnuItm "cotizacion_botom", txtQuote, "javascript:createDocument(23);", "Y", False
	ShowPendOlk = True
End If

If myApp.EnableORDR and CardType <> "S" Then
	AddActiveMnuItm "pedidos_botom", txtOrdr, "javascript:createDocument(17);", "Y", False
	ShowPendOlk = True
End If

If myApp.EnableOPOR and CardType = "S" Then
	AddActiveMnuItm "pedidos_botom", txtOpor, "javascript:createDocument(22);", "Y", False
	ShowPendOlk = True
End If

If CardType = "C" Then

	If myApp.EnableODLN Then
		lnk = "javascript:createDocument(15);"
		AddActiveMnuItm "entrega_icon", txtOdln, lnk, "Y", False
		ShowPendOlk = True
	End If

	If myApp.EnableODPIReq Then
		lnk = "javascript:createDocument(203);"
		AddActiveMnuItm "cotizacion_botom", txtODPIReq, lnk, "Y", False
		ShowPendOlk = True
	End If
	
	If myApp.EnableODPIInv Then
		lnk = "javascript:createDocument(204);"
		AddActiveMnuItm "factura_botom", txtODPIInv, lnk, "Y", False
		ShowPendOlk = True
	End If
	
	If myApp.EnableOINV Then
		lnk = "javascript:createDocument(13);"
		AddActiveMnuItm "factura_botom", txtInv, lnk, "Y", False
		ShowPendOlk = True
	End If
	
	If myApp.EnableOINVRes Then
		lnk = "javascript:createDocument(-13);"
		AddActiveMnuItm "factura_res_botom", txtInvRes, lnk, "Y", False
		ShowPendOlk = True
	End If
	
	If myApp.EnableORCT Then
		lnk = "javascript:createPayment();"
		AddActiveMnuItm "recivo_botom", txtRct, lnk, "N", False
		ShowPendOlk = True
	End If
	
	If myApp.EnableCashInv Then
		lnk = "javascript:createDocument(48);"
		AddActiveMnuItm "mix_botom", txtInv & "/" & txtRct, lnk, "Y", False
		ShowPendOlk = True
	End If
End If

Sub AddActiveMnuItm(ByVal img, ByVal str, ByVal lnk, ByVal Lock, ByVal Selected)
	Dim myActiveItm(4)
	myActiveItm(0) = img
	myActiveItm(1) = str
	myActiveItm(2) = lnk
	myActiveItm(3) = Lock
	myActiveItm(4) = Selected
	
	mnuCount = mnuCount + 1
	ReDim Preserve myActiveMnu(mnuCount)
	myActiveMnu(mnuCount) = myActiveItm
End Sub %>
<table border="0" cellpadding="0" width="100%">
	<tr class="GeneralTlt">
		<td colspan="2">
		<table cellpadding="0" cellspacing="0" border="0" width="100%">
			<tr class="GeneralTlt">
				<td style="padding-top: 2px; padding-left: 4px; padding-right: 4px; width: 1px; ">
				<% Select Case CardType
					Case "C" %>
				<img src="ventas/images/icon_supplier.gif" alt="<%=txtClient%>">
				<%	Case "L" %>
				<img src="ventas/images/icon_lead.gif" alt="<%=getactiveClientLngStr("DtxtLead")%>">
				<% Case "S" %>
				<img src="ventas/images/icon_client.gif" alt="<%=getactiveClientLngStr("DtxtSupplier")%>">
				<% End Select%>
				</td>
				<td><%=getactiveClientLngStr("LttlClientOps")%> - <%=rh("CardName")%></td>
				<td align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>" style="padding-top: 2px; padding-left: 4px; padding-right: 4px; width: 1px;">
					<img src="images/country/pic.aspx?filename=<%=Country%>.gif&amp;MaxHeight=18" alt="<%=CountryName%>">
				</td>
				</tr>
				</table>
				
	</tr>
	<tr>
		<td width="100" align="center">
		<nobr><a href="javascript:Start('thumb/?card=<%=Session("username")%>&pop=Y&AddPath=../',529,510,'yes')"><img src='pic.aspx?dbName=<%=Session("olkdb")%>&amp;filename=<%=crdPic%>&amp;MaxSize=120' style="border: 1px solid #A2D1FD"><% If rs("MoreImg") = "Y" Then %><img src="images/plus.gif" border="0" alt="<%=getactiveClientLngStr("DtxtMore")%>"><% End If %></a></nobr>
		</td>
		<td valign="top">
			<table border="0" cellpadding="0" width="100%">
				<tr class="GeneralTbl">
				<% 
				'iFrom = 0
				'If not optOfert Then iFrom = 1
				'For i = iFrom to UBound(myActiveMnu)+myMnuRest
				
				col = 0
				For i = 0 to UBound(myActiveMnu)+myMnuRest
				If i <> 1 or i = 1 and optOfert Then
				mnuItm = myActiveMnu(i)
				col = col + 1 %>
				<td width="33%">
				<div align="center" title="<%=myHTMLEncode(mnuItm(1))%>">
					<table border="0" cellpadding="0" width="224" style="cursor: hand" onclick="javascript:lnkOp<%=i%>.click();">
						<tr>
							<td align="<% If Session("rtl") = "" Then %>left<% Else %>right<% End If %>" background="design/0/images/<%=addRtl%><%=mnuItm(0)%><% If mnuItm(4) Then %>_on<% End If %>.gif" height="22" id="mnuItm<%=i%>"<% If Not mnuItm(4) Then %> onmouseleave="this.background='design/0/images/<%=addRtl%><%=mnuItm(0)%>.gif';" onmouseenter="this.background='design/0/images/<%=addRtl%><%=mnuItm(0)%>_on.gif';"<% End If %>>
							<div style="width: 200px; height: 18px; overflow: hidden;">
		                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a id="lnkOp<%=i%>" class="LinkOper" href="<% If Not mnuItm(4) Then %><% If mnuItm(3) = "N" or mnuItm(3) = "Y" and ListNum > -1 Then %><%=mnuItm(2)%><% Else %>javascript:listNumAlrt();<% End If %><% Else %>#<% End If %>"><%=myHTMLEncode(mnuItm(1))%></a></div></td>
						</tr>
					</table>
				</div>
				</td>
				<% 'If i = 2+iFrom or i = 5+iFrom or i = 8+iFrom Then Response.Write "</tr><tr class=""GeneralTbl"">"
				If col = 3 Then 
					Response.Write "</tr><tr class=""GeneralTbl"">"
					col = 0
				End If
				End If
				Next %>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td class="RepTotalCel" height="4" colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td height="4" colspan="2" style="font-size: 4px;">&nbsp;</td>
	</tr>
</table>
<% 
Dim MyMnuTab()
mnuCount = -1

Sub AddMnuTabItm(ByVal DisplayText, ByVal Link, ByVal Selected)
	Dim myTopItm(2)
	myTopItm(0) = DisplayText
	myTopItm(1) = Link
	myTopItm(2) = Selected
	
	mnuCount = mnuCount + 1
	ReDim Preserve MyMnuTab(mnuCount)
	MyMnuTab(mnuCount) = myTopItm
End Sub

If Request("open") = "" Then
	If CardType = "S" and myApp.DefClientOPTab = 3 Then
		genData = True
	Else
		Select Case myApp.DefClientOPTab
			Case 3
				pendOlk = True
			Case 5
				openOferts = True
			Case 1
				genData = True
			Case 2
				openAct = True
			Case 4
				openSBO = True
			'Case "maps"
			'	openMaps = True
		End Select
	End If
Else
	Select Case Request("open")
		Case "openedDocs"
			pendOlk = True
		Case "oferts"
			openOferts = True
		Case "genData"
			genData = True
		Case "openAct"
			openAct = True
		Case "openSBO"
			openSBO = True
		Case "maps"
			openMaps = True
		Case "openSO"
			openSO = True
	End Select
End If

AddMnuTabItm getactiveClientLngStr("LtxtGenData"), "doMyLink('activeClient.asp', 'open=genData&CardCode=" & Session("UserName") & "', '');", genData
If myApp.EnableOCLG Then AddMnuTabItm getactiveClientLngStr("DtxtActivities"), "doMyLink('activeClient.asp', 'open=openAct', '');", openAct
If myApp.EnableOOPR Then AddMnuTabItm getactiveClientLngStr("DtxtSO"), "doMyLink('activeClient.asp', 'open=openSO', '');", openSO
If ShowPendOLK Then AddMnuTabItm getactiveClientLngStr("LtxtPendOlk"), "doMyLink('activeClient.asp', 'cCode=" & Session("UserName") & "&open=openedDocs', '');", pendOlk
AddMnuTabItm getactiveClientLngStr("LtxtPendSBO"), "doMyLink('activeClient.asp', 'open=openSBO', '');", openSBO 
If CardType <> "S" Then AddMnuTabItm txtOferts, "doMyLink('activeClient.asp', 'open=oferts', '');", openOferts
If 1 = 2 Then AddMnuTabItm getactiveClientLngStr("LtxtMaps"), "doMyLink('activeClient.asp', 'open=maps', '');", openMaps

%>
<script type="text/javascript">
function tabMenuOver(mnu, i)
{
	mnu.className='ActiveMenuBtnHover';
	document.getElementById('imgTopL' + i).src = 'ventas/images/ActiveMenuTopHover<% If Session("rtl") = "" Then %>2<% End If %>.jpg';
	document.getElementById('imgTopR' + i).src = 'ventas/images/ActiveMenuTopHover<% If Session("rtl") <> "" Then %>2<% End If %>.jpg';
}
function tabMenuOut(mnu, i)
{
	mnu.className='ActiveMenuBtn';
	document.getElementById('imgTopL' + i).src = 'ventas/images/ActiveMenuTopImg<% If Session("rtl") = "" Then %>2<% End If %>.jpg';
	document.getElementById('imgTopR' + i).src = 'ventas/images/ActiveMenuTopImg<% If Session("rtl") <> "" Then %>2<% End If %>.jpg';
}
</script>
<table cellpadding="0" cellspacing="0" border="0">
	<tr>
		<% For i = 0 to UBound(MyMnuTab)
		strText = MyMnuTab(i)(0)
		strLink = MyMnuTab(i)(1)
		isSelected = MyMnuTab(i)(2) %>
		<td><img src="ventas/images/ActiveMenuTop<% If isSelected Then %>Hover<% Else %>Img<% End If %><% If Session("rtl") = "" Then %>2<% End If %>.jpg" id="imgTopL<%=i%>" border="0"></td>
		<td <% If not isSelected Then %>class="ActiveMenuBtn" onmouseover="tabMenuOver(this, <%=i%>);" onmouseout="tabMenuOut(this, <%=i%>);" onclick="<%=strLink%>"<% Else %>class="ActiveMenuBtnHover"<% End If %>><%=strText%></td>
		<td><img src="ventas/images/ActiveMenuTop<% If isSelected Then %>Hover<% Else %>Img<% End If %><% If Session("rtl") <> "" Then %>2<% End If %>.jpg" id="imgTopR<%=i%>" border="0"></td>
		<% Next %>
	</tr>
</table>


<% If pendOlk then %><!--#include file="ventas/ventasx.asp" -->
<% elseif openOferts then %>
<!--#include file="ventas/ofertsX.asp" -->
<% ElseIf openAct Then %>
<!--#include file="ventas/activityX.asp" -->
<% ElseIf openSO Then %>
<!--#include file="addSO/searchOpenedSO.asp"-->
<% ElseIf openSBO Then %>
<!--#include file="portal/openDocs.asp" -->
<% ElseIf genData Then %>
<% addLngPathStr = "addCard/" %>
<!--#include file="addCard/crdConfDetail.asp"-->
<% ElseIf openMaps Then %>
<!--#include file="ventas/maps.asp"-->
<% End If %>
<!--#include file="agentBottom.asp"-->
<!--#include file="clientInc.asp"-->
<% Select Case userType
Case "C"
	user = Session("UserName")
	MainDoc = "clientes" %><!--#include file="clientTop.asp"-->
<% 
If (Session("UserName") = "-Anon-") Then Response.Redirect "default.asp"
Case "V"
	user = Session("vendid")
	MainDoc = "ventas" %><!--#include file="agentTop.asp"-->
<%
End Select
addLngPathStr = "" %>
<!--#include file="lang/messageDetail.asp" -->
<%

set rs = Server.CreateObject("ADODB.recordset")
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetMessage" & Session("ID")
cmd.Parameters.Refresh()
cmd("@olkuser") = user
cmd("@olklog") = Request("olklog")
cmd.execute()

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetMessageDetail" & Session("ID")
cmd.Parameters.Refresh()
cmd("@OlkLog") = Request("olklog")
cmd("@User") = user
set rs = cmd.execute()


imgStr = ""
altStr = ""
Select Case rs("OLKUFromType")
	Case "S"
		imgStr = "alert"
		altStr = getmessageDetailLngStr("DtxtAlert")
	Case "C"
		Select Case rs("CardType")
			Case "C"
				imgStr = "supplier"
				altStr = txtClient
			Case "S"
				imgStr = "client"
				altStr = getmessageDetailLngStr("DtxtSupplier")
			Case "L"
				imgStr = "lead"
				altStr = getmessageDetailLngStr("DtxtLead")
		End Select
	Case "V"
		imgStr = "agent"
		altStr = txtAgent
	Case "B"
		imgStr = "system"
		altStr = getmessageDetailLngStr("DtxtSystem")
	Case "E"
		imgStr = "alert_red"
		altStr = getmessageDetailLngStr("DtxtError")
End Select

If rs("OlkLinkType") = "L" or rs("OlkLinkType") = "D" or rs("OlkLinkType") = "C" or rs("OlkLinkType") = "I" or rs("OlkLinkType") = "A" Then
    If rs("OlkLinkType") = "L" Then
    	DocType = -2
	    Object = rs("OlkLinkObject")
	Else
		Object = rs("OlkLinkObject")
		DocType = Object
	End If
    Select Case Object
    	Case 2
    		btnDesc = txtClient '"Cliente"
    		Action = "addCard/crdConfDetailOpen.asp"
    	Case 4
    		btnDesc = "Articulo"
    		Action = "addItem/itmConfDetail.asp"
    	Case 13, 15, 17, 22, 23, 112, 203, 204
    		Action = "cxcDocDetailOpen.asp"
    		Select Case Object
    			Case 13
		    		btnDesc = txtInv 
		    	Case 15
		    		btnDesc = txtOdln
		    	Case 17
		    		btnDesc = txtOrdr 
		    	Case 23
		    		btnDesc = txtQuote 
		    	Case 112
		    		btnDesc = getmessageDetailLngStr("DtxtDraft")
		    	Case 203
		    		btnDesc = txtODPIReq 
		    	Case 204
		    		btnDesc = txtODPIInv
		    	Case 22
		    		btnDesc = txtOpor
		    End Select
    	Case 24, 140
    		Select Case Object
    			Case 24
		    		btnDesc = txtRct '"Recibo"
		    	Case 140
		    		btnDesc = getmessageDetailLngStr("DtxtDraft")
		    End Select
    		Action = "cxcRctDetailOpen.asp"
    	Case 33
    		btnDesc = getmessageDetailLngStr("DtxtActivity")
    		Action = "addActivity/activityConfDetail.asp"
    	Case 97
    		btnDesc = getmessageDetailLngStr("DtxtSO")
    		Action = "addSO/SOConfDetail.asp"
    End Select
End If
%>
<table border="0" cellpadding="0" width="100%">
	<% If tblCustTtl = "" Then %>
	<tr class="TablasTituloSec">
		<td colspan="3" id="tdMyTtl">
		<table cellpadding="0" cellspacing="1" width="100%" border="0">
			<tr class="MsgTlt">
				<td width="20"><img border="0" src="ventas/images/icon_<%=imgStr%>.gif" alt="<%=Server.HTMLEncode(altStr)%>"></td>
				<td>&nbsp;<%=Replace(getmessageDetailLngStr("LttlMsgTitle"), "{0}", Request("OlkLog"))%></td>
			</tr>
		</table>
		</td>
	</tr>
	<% Else %>
	<tr>
		<td colspan="3">
			<table cellpadding="0" cellspacing="0" width="100%" border="0">
			<tr>
				<td>
				<% strTitle = "<img border=""0"" src=""ventas/images/icon_" & imgStr & ".gif"" alt=""" & Server.HTMLEncode(altStr) & """>" & Replace(getmessageDetailLngStr("LttlMsgTitle"), "{0}", Request("OlkLog")) %>
				<%=Replace(Replace(tblCustTtl, "{txtTitle}", strTitle), "{AddPath}", "")%>
				</td>
			</tr>
			</table>
		</td>
	</tr>
	<% End If %>
	<tr class="CanastaTblResaltada">
		<td width="49%" colspan="2">
		<p align="center"><%=getmessageDetailLngStr("DtxtDate")%></td>
		<td width="50%">
		<p align="center"><% If rs("OlkUFromType") = "V" or rs("OlkUFromType") = "C" Then %><%=getmessageDetailLngStr("LtxtFrom")%><% Else %><%=getmessageDetailLngStr("DtxtType")%><% End If %></td>
	</tr>
	<tr class="CanastaTbl">
		<td width="30"><nobr><% If RS("olkUrgent") = "Y" Then %>
		<img border="0" src="images/mail_icon_urgent.gif" width="13" height="12"><% End If %>
		<% If CStr(Trim(RS("OlkStatus"))) = "N" Then OImg = "new" Else OImg = "open" %><img border="0" src="images/mail_icon_<%=OImg%>.gif"></nobr>
		</td>
		<td>
		<p align="center"><%=FormatDate(RS("Date"), True)%>&nbsp;<%=rs("Time")%></td>
		<td>
		<p align="center"><% Select Case rs("olkufromType")
				Case "S"
					Response.Write getmessageDetailLngStr("DtxtAlert")
				Case "E"
					Response.Write getmessageDetailLngStr("DtxtError")
				Case "B"
					Response.Write getmessageDetailLngStr("DtxtSystem")
				Case Else
					Response.Write RS("olkUFromName")
		End Select %></td>
	</tr>
	<tr class="CanastaTblResaltada">
		<td colspan="3">
		<p align="center"><%=getmessageDetailLngStr("DtxtSubject")%></td>
	</tr>
	<tr class="CanastaTbl">
		<td colspan="3">
		<p align="center"><%=RS("OlkSubject")%>&nbsp;</td>
	</tr>
	<tr>
		<td colspan="3">
		<table border="0" cellpadding="0" width="100%" cellspacing="1">
				<tr class="CanastaTbl">
					<td>
					<p align="center"><textarea readonly rows="11" name="S1" cols="58" style="width: 100%"><%
					If rs("OlkUFromType") = "E" Then %>
					<%=getmessageDetailLngStr("DtxtCode")%>: <%=rs("ErrCode") & VbNewLine %>
					<%=getmessageDetailLngStr("DtxtError")%>: <%=rs("ErrMessage") & VbNewLine & VbNewLine %>
					<% If rs("Status") = "S" Then %>
					<%=Replace(Replace(getmessageDetailLngStr("LtxtUpdate"), "{0}", btnDesc), "{1}", rs("DocNum"))%><%=VbNewLine%>
					<%=getmessageDetailLngStr("DtxtDate")%>: <%=rs("EndDate")%><%=VbNewLine%>
					<%=getmessageDetailLngStr("DtxtHour")%>: <%=rs("EndTime")%><%=VbNewLine%>
					<% End If %>
					<% Else %>
					<%=myHTMLEncode(RS("OlkMsg"))%>
					<% End If %></textarea></td>
				</tr>
				<tr class="CanastaTbl">
					<td>
					<p align="center">
					<% If rs("OlkLinkType") = "O" Then
					If userType = "V" Then
						set ro = Server.CreateObject("ADODB.RecordSet")
						set cmd = Server.CreateObject("ADODB.Command")
						cmd.ActiveConnection = connCommon
						cmd.CommandType = &H0004
						cmd.CommandText = "DBOLKGetOfferLinkData" & Session("ID")
						cmd.Parameters.Refresh()
						cmd("@OlkLog") = Request("OlkLog")
						set ro = cmd.execute()
						CardCode = ro("CardCode")
						ItemCode = ro("ItemCode")
						ro.close
						set ro = nothing
					End If %>
			    <input type="button" class="btnOffer" value="<%=Replace(getmessageDetailLngStr("LtxtView"), "{0}", Server.HTMLEncode(txtOfert))%>" onClick="javascript:<% If userType = "C" Then %>window.location.href='ofertHistory.asp?ofertIndex=<%=rs("olkLink")%>';<% Else %>window.location.href='ofertsMan.asp?CardCodeFrom=<%=CardCode%>&CardCodeTo=<%=CardCode%>&ItemCodeFrom=<%=ItemCode%>&ItemCodeTo=<%=ItemCode%>';<% End If %>">
			    <% ElseIf rs("OlkLinkType") = "L" or rs("OlkLinkType") = "D" or rs("OlkLinkType") = "C" or rs("OlkLinkType") = "I" or rs("OlkLinkType") = "A" Then
			    olkLink = rs("olkLink")
			    If rs("OLKUFromType") = "E" and rs("Status") = "S" Then 
			    	olkLink = rs("ObjectCode")
			    	DocType = Object
			    End If
			    If rs("OLKUFromType") = "E" and rs("Status") <> "S" Then %>
			    <input type="button" class="btnLink" value="<%=Replace(getmessageDetailLngStr("LtxtGoTo"), "{0}", myHTMLEncode(btnDesc))%>" onClick="javascript:Open();">
			    <% Else %>
			    <input type="button" class="btnView" value="<%=Replace(getmessageDetailLngStr("LtxtView"), "{0}", myHTMLEncode(btnDesc))%>" onClick="javascript:GoLogView('<%=Action%>', '<%=olkLink%>');">
			    <% End If %>
			    <% End If %></td>
				</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td colspan="3">
		<table border="0" cellpadding="0" width="100%">
			<tr class="CanastaTblResaltada">
				<td align="center">
				<a href="javascript:<% If userType = "C" Then %>doMyLink('messageReturn.asp', 'cmd=newMsg&to=<%=myHTMLEncode(RS("olkufrom"))%>&tot=<%=RS("olkufromtype")%>&OlkLog=<%=rs("OlkLog")%>', '');<% Else %>window.location.href='newMessage.asp';<% End If %>">
				<img border="0" src="design/<%=SelDes%>/images/msg_icon_newmsg.gif" width="47" height="40" alt="<%=getmessageDetailLngStr("LaltNewMsg")%>"></a></td>
				<td align="center">
				<a href="javascript:<% If rs("olkufromtype") <> "S" and rs("olkufromtype") <> "E" and rs("olkufromtype") <> "B" Then %>doMyLink('messageReturn.asp', 'cmd=reply&to=<%=myHTMLEncode(RS("olkufrom"))%>&tot=<%=RS("olkufromtype")%>&OlkLog=<%=rs("OlkLog")%>', '');<% Else %>alert('<%=getmessageDetailLngStr("LtxtSystemReplyErr")%>');<% End If %>">
				<img border="0" src="design/<%=SelDes%>/images/msg_icon_responder.gif" width="47" height="40" alt="<%=getmessageDetailLngStr("LtxtReply")%>"></a></td>
				<td align="center">
				<a href='messages/updateMsgStatus.asp?olklog=<%=Request("olklog")%>&amp;status=N&amp;pop=N&amp;AddPath=<%=Request("AddPath")%>'>
				<img border="0" src="design/<%=SelDes%>/images/msg_icon_archivar.gif" width="47" height="40" alt="<%=getmessageDetailLngStr("LtxtFile")%>"></a></td>
				<td align="center">
				<a href="javascript:if(confirm('<%=getmessageDetailLngStr("LtxtConfDelMsg")%>'))window.location.href='messages/msgDelete.asp?olklog=<%=Request("olklog")%>&pop=N&AddPath=<%=Request("AddPath")%>'">
				<img border="0" src="design/<%=SelDes%>/images/msg_icon_borrar.gif" width="47" height="40" alt="<%=getmessageDetailLngStr("LtxtDel")%>"></a></td>
			</tr>
		</table>
		</td>
	</tr>
</table>
<form target="_blank" method="post" name="viewLogNum">
<input type="hidden" name="DocEntry" value="">
<input type="hidden" name="DocType" value="<%=DocType%>">
<input type="hidden" name="CardCode" value="">
<input type="hidden" name="ItemCode" value="">
<input type="hidden" name="pop" value="Y">
<input type="hidden" name="AddPath" value="">
</form>
<script language="javascript">
function GoLogView(Action, LogNum) 
{
	document.viewLogNum.action = Action
	<% Select Case rs("OlkLinkType")
		Case "I" %>
	document.viewLogNum.ItemCode.value = LogNum 
	<%	Case "C" %>
	document.viewLogNum.CardCode.value = LogNum 
	<%	Case Else %>
	document.viewLogNum.DocEntry.value = LogNum 
	<% End Select %>
	if (Action != 'addItem/itmConfDetail.asp' && Action != 'addCard/crdConfDetailOpen.asp') { document.viewLogNum.AddPath.value = ''; }
	else { document.viewLogNum.AddPath.value = '../'; }
	document.viewLogNum.submit() 
}

function Open()
{
	<% 
	Select Case Object
		Case 2
			frmAction = "ventas/goCard.asp"
		Case 4
			frmAction = "ventas/goItem.asp"
		Case 24
			frmAction = "payments/go.asp"
			goDoc = True
		Case 33
			frmAction = "addActivity/goActivity.asp"
		Case 13,15,17,23
			frmAction = "ventas/go.asp"
			goDoc = True
	End Select 
	
	If goDoc Then
	%>
	doMyLink('<%=frmAction%>', 'doc=<%=rs("olkLink")%>&payDoc=&cl=<%=rs("CardCode")%>&status=R', '');
	<% ElseIf Object = 33 Then %>
	doMyLink('<%=frmAction%>', 'LogNum=<%=rs("olkLink")%>&Card=<%=rs("CardCode")%>&status=R', '');
	<% Else %>
	doMyLink('<%=frmAction%>', 'LogNum=<%=rs("olkLink")%>&status=R', '');
	<% End If %>
	window.close();
}
</script>

<% If setCustTtl and userType = "C" Then %>
<script language="javascript" src="setTltBg.js.asp?custTtlBgL=<%=custTtlBgL%>&amp;custTtlBgM=<%=custTtlBgM%>&amp;AddPath=../"></script>
<script language="javascript">setTtlBg(false);</script>
<% End If %>
<% Select Case userType
Case "C" %><!--#include file="clientBottom.asp"-->
<% Case "V" %><!--#include file="agentBottom.asp"-->
<% End Select %>
<!--#include file="clientInc.asp"-->
<!--#include file="agentTop.asp"-->
<% If userType <> "V" Then Response.Redirect "default.asp" %>
<% 
If Not (EnableClientActivation and myAut.HasAuthorization(89)) Then Response.Redirect "unauthorized.asp"
addLngPathStr = "" %>
<!--#include file="lang/activation.asp" -->
<SCRIPT LANGUAGE="JavaScript">

<!-- Begin
function Start(page) {
OpenWin = this.open(page, "confirmDocs", "toolbar=no,menubar=no,location=no,scrollbars=no,resizable=no, height=378,width=420");
OpenWin.focus()
}
// End -->
</SCRIPT>
<%
varxx = 0
           set rs = Server.CreateObject("ADODB.recordset")

           set rf = Server.CreateObject("ADODB.recordset")
If Request("new") = "Y" Then Session("confStr") = ""

set rd = Server.CreateObject("ADODB.RecordSet")
sql = "select SlpCode, IsNull(SlpName, '') SlpName from OSLP order by 2 asc"
set rd = conn.execute(sql)

sql = "select T0.CardCode, IsNull(T0.CardName, '') CardName, IsNull(T1.GroupName, '') GroupName, IsNull(T2.Name, '') Country, T0.CreateDate CreateDate, T0.DocEntry " & _
"from OCRD T0 " & _
"inner join OCRG T1 on T1.GroupCode = T0.GroupCode " & _
"left outer join OCRY T2 on T2.Code = T0.Country " & _
"inner join OLKClientsAccess T3 on T3.CardCode = T0.CardCode " & _
"where T3.Status = 'P' order by T0.CardCode asc "
rs.open sql, conn, 3, 1 
%>
<script language="javascript">
function GoLogView(CardCode) {
document.viewLogNum.CardCode.value = CardCode 
document.viewLogNum.submit() }

function goViewFlow(LogNum)
{
	OpenWin = this.open("flowAlertView.asp?LogNum=" + LogNum + "&pop=Y", "CtrlWindow", "toolbar=no,menubar=no,location=no,scrollbars=yes,resizable=no,width=550,height=400,top="+wint+",left="+winl);
}
</script>
<form target="_blank" method="post" name="viewLogNum" action="addCard/crdConfDetailOpen.asp">
<input type="hidden" name="CardCode" value="">
<input type="hidden" name="DocType" value="2">
<input type="hidden" name="pop" value="Y">
<input type="hidden" name="AddPath" value="../">
</form>
<script language="javascript">
function valFrm()
{
	found = false;
	<% If rs.recordcount > 1 Then %>
	myStatus = document.frmActiveAcct.Status;
	for (var i = 0;i<myStatus.length;i++)
	{
		if (myStatus(i).value == 'A' || myStatus(i).value == 'R') found = true;
		<% If myApp.AnRegConfAsignSLP Then %>
		if (myStatus(i).value == 'A' && document.frmActiveAcct.Slp(i).value == '-1')
		{
			alert('<%=getactivationLngStr("LtxtValActiveAgent")%>'.replace('{0}', '<%=txtAgents%>'));
			return false;
		}
		<% End If %>
		<% If myApp.AnRegConfRejNote Then %>
		if (myStatus(i).value == 'R' && document.frmActiveAcct.note(i).value == '')
		{
			alert('<%=getactivationLngStr("LtxtValRejReason")%>');
			return false;
		}
		<% End If %>
	}
	<% Else %>
	if (document.frmActiveAcct.Status.value == 'A' || document.frmActiveAcct.Status.value == 'R') found = true;
	<% If myApp.AnRegConfAsignSLP Then %>
	if (document.frmActiveAcct.Status.value == 'A' && document.frmActiveAcct.Slp.value == '-1')
	{
		alert('<%=getactivationLngStr("LtxtValActiveAgent")%>'.replace('{0}', '<%=txtAgent%>'));
		return false;
	}
	<% End If %>	
	<% If myApp.AnRegConfRejNote Then %>
	if (document.frmActiveAcct.Status.value == 'R' && document.frmActiveAcct.note.value == '')
	{
		alert('<%=getactivationLngStr("LtxtValRejReason")%>');
		return false;
	}
	<% End If %>
	<% End If %>	
	if (!found)
	{
		alert('<%=getactivationLngStr("LtxtValStatAcct")%>');
	}
	return found;
}
function changeStatus(DocEntry, s)
{
	if (s == 'A') { document.getElementById('Slp' + DocEntry).disabled = false; }
	else { document.getElementById('Slp' + DocEntry).disabled = true; }
	if (s == 'R') { document.getElementById('btnEdit' + DocEntry).disabled = false; }
	else { document.getElementById('btnEdit' + DocEntry).disabled = true; }
}
var uNote;
function editReason(note)
{
	uNote = note;
	OpenWin = this.open("", "Reason", "toolbar=no,menubar=no,location=no,scrollbars=no,resizable=no, width=400,height=400");
	doMyLink('ventas/activationReason.asp', 'note=' + note.value + '&pop=Y&AddPath=../', 'Reason');
	OpenWin.focus()
}
function setReason(note)
{
	uNote.value = note;
}
</script>
<form method="POST" action="activationSubmit.asp" name="frmActiveAcct" onsubmit="return valFrm();">
<table border="0" cellpadding="0" width="100%" id="table1">
	<tr class="GeneralTlt">
		<td>&nbsp;<%=getactivationLngStr("LttlAnonRegAct")%></td>
	</tr>
	<tr class="GeneralTblBold2">
		<td><% If 1 = 2 Then %><%=getactivationLngStr("DtxtClients")%><% Else %><%=txtClients%><% End If %></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table3">
			<tr class="GeneralTblBold2">
				<td align="center" width="15">&nbsp;</td>
				<td align="center"><%=getactivationLngStr("DtxtCode")%></td>
				<td align="center"><%=getactivationLngStr("DtxtName")%></td>
				<td align="center"><%=getactivationLngStr("DtxtGroup")%></td>
				<td align="center"><%=getactivationLngStr("DtxtCountry")%></td>
				<td align="center"><%=getactivationLngStr("DtxtDate")%></td>
				<td align="center"><%=getactivationLngStr("DtxtState")%></td>
				<td align="center"><%=getactivationLngStr("LtxtAsign")%>&nbsp;<%=txtAgent%></td>
				<td align="center"><%=getactivationLngStr("LtxtRejReas")%></td>
			</tr>
			<%
			if not rs.eof then
			do while not rs.eof
			ShowSubmit = True %>
			<tr class="GeneralTbl">
				<td width="15">
				<a href="javascript:GoLogView('<%=rs("CardCode")%>')">
				<img border="0" src="design/0/images/<%=Session("rtl")%>felcahSelect.gif" width="15" height="13"></a></td>
				<td><%=rs("CardCode")%>&nbsp;</td>
				<td><%=rs("CardName")%>&nbsp;</td>
				<td><%=rs("GroupName")%>&nbsp;</td>
				<td><%=rs("Country")%>&nbsp;</td>
				<td><%=FormatDate(rs("CreateDate"), True)%>&nbsp;</td>
				<td>
				<select size="1" id="Status" name="Status<%=rs("DocEntry")%>" onchange="javascript:changeStatus(<%=rs("DocEntry")%>, this.value);">
				<option value="P"><%=getactivationLngStr("LtxtPending")%></option>
				<option value="A"><%=getactivationLngStr("LtxtActivate")%></option>
				<option value="R"><%=getactivationLngStr("LtxtReject")%></option>
				</select></td>
				<td>
				<p align="center">
				<select disabled size="1" id="Slp" name="Slp<%=rs("DocEntry")%>">
				<% rd.movefirst
				do while not rd.eof %>
				<option <% If rd(0) = -1 Then %>selected<% End If %> value="<%=rd(0)%>"><%=myHTMLEncode(rd(1))%></option>
				<% rd.movenext
				loop %>
				</select></td>
				<td>
				<p align="center">
				<input disabled type="button" value="<%=getactivationLngStr("DtxtEdit")%>" name="btnEdit<%=rs("DocEntry")%>" onclick="javascript:editReason(document.frmActiveAcct.note<%=rs("DocEntry")%>)">
				<input type="hidden" id="note" name="note<%=rs("DocEntry")%>" value=""></td>
			</tr>
			<% rs.movenext
			loop
			Else %>
			<tr class="GeneralTblBold2">
				<td colspan="9">
				<p align="center"><%=getactivationLngStr("DtxtNoData")%></td>
			</tr>
			<% End If %>
		</table>
		</td>
	</tr>
	<tr class="GeneralTbl">
		<td>
		<% If ShowSubmit Then %><input type="submit" value="<%=getactivationLngStr("DtxtAccept")%>" name="Aceptar" style="float: right"><% End If %></td>
	</tr>
</table>
<input type="hidden" name="cmd" value="activationSubmit">
<input type="hidden" name="submit" value="Y">
</form>
<!--#include file="agentBottom.asp"-->
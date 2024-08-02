<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="../chkLogin.asp" -->
<!--#include file="lang/cards.asp" -->
<!--#include file="../myHTMLEncode.asp"-->
<html <% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>>

<%
           set rs = Server.CreateObject("ADODB.RecordSet")
           If Session("PayCart") Then sqlAdd = "OIR"
           sql = "select T0.CreditCard, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRC', 'CardName', T0.CreditCard, CardName) CardName, " & _
           		 "Case When (select EnbSgmnAct from CINF) = 'Y' Then OLKCommon.dbo.DBOLKGetSegmentAccount" & Session("ID") & "(IsNull(T2." & sqlAdd & "AcctCode,T0.AcctCode)) Else IsNull(T2." & sqlAdd & "AcctCode,T0.AcctCode) End AcctDisp, " & _
           		 "IsNull(T2." & sqlAdd & "AcctCode,T0.AcctCode) AcctCode, CrTypeCode, IsNull(CrTypeName, '') CrTypeName, MinCredit, MinToPay, MaxValid, InstalMent " & _
           		 "from ocrc T0 " & _
           		 "left outer join OCRP T1 on T1.CreditCard = T0.CreditCard " & _
           		 "left outer join OLKBranchsAcctCredit T2 on T2.branchIndex = " & Session("branch") & " and T2.CreditCard = T0.CreditCard"
           set rs = conn.execute(sql)
           %>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>
<%=getcardsLngStr("LttlCredCardSel")%></title>
<link rel="stylesheet" type="text/css" href="../design/0/style/stylePopUp.css">

</head>

<script language="javascript">
function setCard(CreditCard, CardName, CrTypeCode, CrTypeName, AcctCode, MinCredit, MinToPay, MaxValid, InstalMent, AcctDisp)
{
	opener.setCard(CreditCard, CardName, CrTypeCode, CrTypeName, AcctCode, MinCredit, MinToPay, MaxValid, InstalMent, AcctDisp);
	opener.clearWin();
	window.close();
}
</script>
<body marginwidth="0" marginheight="0" topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" bgcolor="#EDF5FE">
<table border="0" cellpadding="0" width="100%" id="table1">
	<tr>
		<td colspan="5" class="GeneralTlt"><%=getcardsLngStr("LttlSelCard")%></td>
	</tr>
	<tr>
		<td class="GeneralTblBold2">
		<%=getcardsLngStr("LtxtCard")%></td>
		<td class="GeneralTblBold2">
		<%=myHTMLDecode(getcardsLngStr("LtxtMinCred"))%></td>
		<td class="GeneralTblBold2">
		<%=myHTMLDecode(getcardsLngStr("LtxtMinPymnt"))%></td>
		<td class="GeneralTblBold2">
		<%=myHTMLDecode(getcardsLngStr("LtxtMaxAut"))%></td>
		<td class="GeneralTblBold2">
		<%=getcardsLngStr("LtxtPymnts")%></td>
	</tr>
	<% If not rs.eof then 
	while not rs.eof %>
	<tr class="GeneralTbl" onclick="javascript:setCard('<%=myHTMLEncode(rs("CreditCard"))%>','<%=myHTMLEncode(rs("CardName"))%>','<%=myHTMLEncode(rs("CrTypeCode"))%>','<%=myHTMLEncode(rs("CrTypeName"))%>','<%=rs("AcctCode")%>','<%=rs("MinCredit")%>','<%=rs("MinToPay")%>','<%=rs("MaxValid")%>','<%=rs("InstalMent")%>','<%=rs("AcctDisp")%>');" style="cursor: hand" onmouseover="javascript:this.className='GeneralTblHigh'" onmouseout="javascript:this.className='GeneralTbl'">
		<td>
		<%=rs("CardName")%></td>
		<td>
		<%=rs("MinCredit")%>&nbsp;</td>
		<td>
		<%=rs("MinToPay")%>&nbsp;</td>
		<td>
		<%=rs("MaxValid")%>&nbsp;</td>
		<td>
		<% Select Case rs("InstalMent")
		Case "N"
		response.write getcardsLngStr("DtxtNo")
		Case "Y"
		response.write getcardsLngStr("DtxtYes")
		End Select %>&nbsp;</td>
	</tr>
	<% rs.movenext
	wend 
	else %>
	<tr class="GeneralTbl">
		<td colspan="5">
		<%=getcardsLngStr("LtxtNoCard")%></td>
	</tr>
	<% end if %>
	</table>

</body>
<% conn.close %>
</html>
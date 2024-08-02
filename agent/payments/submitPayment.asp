<% addLngPathStr = "payments/" %>
<!--#include file="lang/submitPayment.asp" -->
<%
If Session("PayRetVal") <> "" Then RetVal = Session("PayRetVal") Else RetVal = Session("ConfPayRetVal")
If Request("Confirm") = "N" Then
	If myAut.GetObjectProperty(24, "C") Then
		goStatus = "H"
	Else
		goStatus = "C"
	End If
Else
	goStatus = "H"
End If

rctDesc = txtRct

sql = "select Draft, (select ViewObjPrint from OLKDocConf where ObjectCode = 24) ViewObjPrint from R3_ObsCommon..TLOG where LogNum = " & RetVal
set rs = conn.execute(sql)
If rs("Draft") = "Y" Then 
	objRctCode = 140
	rctDesc = rctDesc & " (" & getsubmitPaymentLngStr("DtxtDraft") & ")"
Else 
	objRctCode = 24
End If
ViewObjPrint = rs("ViewObjPrint") = "Y"

set mySubmit = new SubmitControl
mySubmit.EnableRunInBackground = True
mySubmit.LogNum = RetVal 
mySubmit.LogNumID = "PayRetVal"
mySubmit.TransactionOkMessage = Replace(getsubmitPaymentLngStr("LtxtConfAddPay"), "{1}", rctDesc)
mySubmit.EndButtonDescription = Replace(getsubmitPaymentLngStr("LtxtCreateNewPay"), "{0}", txtRct)
mySubmit.EndButtonFunction = "window.location.href='payments/newDocGo.asp?AddPath=../';"
If ViewObjPrint  Then 
	mySubmit.SecondButtonDescription = getsubmitPaymentLngStr("LtxtView") & " " & txtRct
	mySubmit.SecondButtonFunction = "viewDetails('{0}');"
End If
mySubmit.RunInBackgroundRedir = "agentPaymentConfirm.asp?RetVal=" & Session("ItmRetVal") & "&Confirm=Y&bg=Y"
mySubmit.GenerateSubmit %>
<script type="text/javascript">
function Start(page) 
{
	OpenWin = this.open(page, 'objDetails', 'toolbar=no,menubar=no,location=no,scrollbars=yes,resizable=yes, height=600,width=800');
	OpenWin.focus()
}

function viewDetails(payCode)
{
	Start('');
	doMyLink('cxcRctDetailOpen.asp', 'DocType=<%=objRctCode%>&DocEntry=' + payCode + '&pop=Y&AddPath=../', 'objDetails');
}
</script>

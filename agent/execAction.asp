<!--#include file="clientInc.asp"-->
<% Select Case userType
Case "C" %><!--#include file="clientTop.asp"-->
<% Case "V" %><!--#include file="agentTop.asp"-->
<% End Select %>
<% addLngPathStr = "" %>
<!--#include file="lang/execAction.asp" -->
<%


If Session("RepLogNum") = "" Then

	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKExecuteAction" & Session("ID")
	cmd.Parameters.Refresh
	cmd("@ActionID") = Request("ID")
	cmd("@Entry") = Request("Entry")
	cmd("@UserSign") = Session("vendid")
	
	Select Case CInt(Request("ID"))
		Case 1, 7 'Convert Quote to Order
			cmd("@Series") = Request("Series")
		Case 2, 3, 6
			cmd("@ObjectCode") = Request("ObjectCode")
	End Select
	
	cmd.Execute()

	LogNum = cmd("@LogNum")

	Session("RepLogNum") = LogNum
Else
	LogNum = Session("RepLogNum")
End If

ID = CInt(Request("ID"))

set mySubmit = new SubmitControl
mySubmit.EnableRunInBackground = False
mySubmit.LogNum = LogNum
mySubmit.EndButtonDescription = getexecActionLngStr("LtxtBackToRep")
mySubmit.EndButtonFunction = "reloadRep();"

Select Case ID
	Case 0
		mySubmit.TransactionOkMessage = getexecActionLngStr("LtxtOrdrAprove")
	Case 1
		mySubmit.TransactionOkMessage = getexecActionLngStr("LtxtOrderConv")
	Case 7
		mySubmit.TransactionOkMessage = getexecActionLngStr("LtxtInvConv")
	Case 2
		mySubmit.TransactionOkMessage = getexecActionLngStr("LtxtObjCanceled")
	Case 3
		mySubmit.TransactionOkMessage = getexecActionLngStr("LtxtObjClosed")
	Case 6
		mySubmit.TransactionOkMessage = getexecActionLngStr("LtxtObjDel")
End Select

If ID = 0 or ID = 1 or ID = 7 Then
	viewStr = txtOrdr
	viewObj = 17
	
	If ID = 7 Then 
		viewStr = txtInv
		viewObj = 13
	End If

	mySubmit.SecondButtonDescription = getexecActionLngStr("LtxtView") & " " & viewStr
	mySubmit.SecondButtonFunction = "openDoc(" & viewObj & ", {0});"
End If

mySubmit.GenerateSubmit

%>
<script type="text/javascript">
<% If ID = 0 or ID = 1 or ID = 7 Then %>
function openDoc(objCode, docNum)
{
	document.viewLogNum.DocType.value = <% If ID = 7 Then %>13<% Else %>17<% End If %>;
	document.viewLogNum.DocEntry.value = docNum;
	document.viewLogNum.submit();
}
<% End If %>
function reloadRep()
{
	doMyLink('report.asp', '<%=Replace(Replace(Request("retVal"), "{y}", "&"), "{i}", "=")%>', '');
}
</script>
<% If ID = 0 or ID = 1 or ID = 7 Then %>
<form target="_blank" method="post" name="viewLogNum" action="cxcDocDetailOpen.asp">
<input type="hidden" name="DocEntry" value="">
<input type="hidden" name="DocType" value="<%=DocType%>">
<input type="hidden" name="pop" value="Y">
<input type="hidden" name="isEntry" value="N">
</form>
<% End If %>
<% Select Case userType
Case "C" %><!--#include file="clientBottom.asp"-->
<% Case "V" %><!--#include file="agentBottom.asp"-->
<% End Select %>
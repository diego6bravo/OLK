<!--#include file="clientInc.asp"-->
<!--#include file="agentTop.asp"-->
<% If userType <> "V" Then Response.Redirect "default.asp" %>
<% 
If Not myApp.EnableOOPR Then Response.Redirect "unauthorized.asp"
addLngPathStr = "" %>
<!--#include file="lang/SOSubmit.asp" -->
<% 
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004

If Session("SORetVal") <> "" Then SORetVal = Session("SORetVal") Else SORetVal = Session("ConfSORetVal")
If Request("doSubmit") = "Y" Then 
	set rs = Server.CreateObject("ADODB.RecordSet")
	cmd.CommandText = "DBOLKGetLogStatus"
	cmd("@LogNum") = SORetVal
	set rs = cmd.execute()
	If rs("Status") = "R" Then
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandText = "DBOLKCheckObjectCheckSum" & Session("ID")
		cmd.CommandType = &H0004
		cmd.Parameters.Refresh()
		cmd("@LogNum") = Session("SORetVal")
		cmd.Execute()
	
		If cmd("@IsValid").Value = "N" Then
			Session("RetryRetVal") = Session("SORetVal")
			Session("SORetVal") = ""
			Response.Redirect "crcerror.asp"
		End If

		Session("NotifyAdd") = True
		
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKPostObjectCreation" & Session("ID")
		cmd.Parameters.Refresh
		cmd("@sessiontype") = "A"
		cmd("@transtype") = "A"
		cmd("@object") = 33 
		cmd("@LogNum") = Session("SORetVal")
		cmd("@CurrentSlpCode") = Session("vendid")
		cmd("@Branch") = Session("branch")
		cmd.execute()
	
		cmd.CommandText = "DBOLKExecuteLog"
		cmd.Parameters.Refresh()
		cmd("@LogNum") = Session("SORetVal")
		cmd("@SlpCode") = Session("vendid")
		cmd("@branchIndex") = Session("branch")
		cmd.execute()
		
		If goStatus <> "H" Then
			doSubmitSOWait()
		Else
			SORetVal = Session("SORetVal")
			ShowNewSO()
		End If
	Else
		doSubmitSOWait()
	End If
Else
	ShowNewActivity()
End If

Sub doSubmitSOWait()
	set mySubmit = new SubmitControl
	mySubmit.EnableRunInBackground = True
	mySubmit.LogNum = SORetVal 
	mySubmit.LogNumID = "SORetVal"
	If Request("isUpdate") = "True" Then
		mySubmit.TransactionOkMessage = getSOSubmitLngStr("LtxtConfUpdSO")
	Else
		mySubmit.TransactionOkMessage = getSOSubmitLngStr("LtxtConfAddSO")
		mySubmit.EndButtonDescription = getSOSubmitLngStr("LtxtCreateNewSO")
		mySubmit.EndButtonFunction = "window.location.href='addSO/goNewSO.asp?AddPath=../';"
		mySubmit.SecondButtonDescription = getSOSubmitLngStr("DtxtView") & " " & getSOSubmitLngStr("DtxtSO")
		mySubmit.SecondButtonFunction = "viewDetails('{0}');"
		mySubmit.RunInBackgroundRedir = "SOSubmit.asp?RetVal=" & Session("SORetVal") & "&Confirm=Y&bg=Y"
	End If
	mySubmit.GenerateSubmit %>
<script type="text/javascript">
function Start(page) 
{
	OpenWin = this.open(page, 'objDetails', 'toolbar=no,menubar=no,location=no,scrollbars=yes,resizable=yes, height=600,width=800');
	OpenWin.focus()
}

function viewDetails(soID)
{
	Start('');
	doMyLink('addSO/SOConfDetail.asp', 'DocType=97&DocEntry=' + soID + '&pop=Y&AddPath=../', 'objDetails');
}
</script>
<% End Sub

Sub ShowNewSO() 
%>
Background
<% End Sub %>
<!--#include file="agentBottom.asp"-->
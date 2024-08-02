<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>

<!--#include file="../myHTMLEncode.asp" -->

<html>

<head>
</head>

<body onload="javascript:doFinish();">
<!--#include file="../itemFunctions.asp" -->
<% 
set rs = Server.CreateObject("ADODB.RecordSet")
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "DBOLKCartAddSFM" & Session("ID")
cmd.Parameters.Refresh()
cmd("@SlpCode") = Session("vendid")

LineNum = Split(Request("LineNum"), ", ")
For i = 0 to UBound(LineNum)
	cmd("@lognum") = Session("RetVal")
	cmd("@Quantity") = CDbl(Request("Quantity" & LineNum(i)))
	cmd("@item") = Request("ItemCode" & LineNum(i))
	cmd("@itemprice") = CDbl(Request("Price" & LineNum(i)))
	cmd("@PriceList") = Session("PriceList")
	cmd("@UserType") = userType
	cmd("@SaleType") = Request("SaleUnit" & LineNum(i))
	cmd("@branchIndex") = Session("branch")
	If myApp.LawsSet = "MX" or myApp.LawsSet = "CL" or myApp.LawsSet = "CR" or myApp.LawsSet = "GT" or myApp.LawsSet = "US" or myApp.LawsSet = "CA" Then
		TaxCode = getItemTaxCode(LineNum(i))
		If TaxCode <> "Disabled" and TaxCode <> "" Then cmd("@TaxCode") = TaxCode
	End If
	cmd.execute()
Next

conn.close %>
<script language="javascript">
opener.location.href='../cart.asp';
window.close();
</script>
</body>

</html>
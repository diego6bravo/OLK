<!--#include file="clientInc.asp"-->
<% Select Case userType
Case "C"
If Session("UserName") = "-Anon-" or myApp.EnableAnonCart and myApp.AnonCartClient = Session("UserName") Then Response.Redirect "clientLogin.asp" 
 %><!--#include file="clientTop.asp"-->
<% 
If (Session("UserName") = "-Anon-" or not optBasket) Then Response.Redirect "default.asp"
Case "V" %><!--#include file="agentTop.asp"-->
<% 
If Not comDocsMenu Then Response.Redirect "unauthorized.asp"
End Select %>
<% addLngPathStr = "" %>
<!--#include file="lang/cartSubmit.asp" -->
<% 

If Request("saveAddress") = "Y" Then
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKSetDocAddress" & Session("ID")
	cmd.Parameters.Refresh
	cmd("@LogNum") = Session("RetVal")
	If Request("BillToCode") <> "" Then cmd("@BillToCode") = Request("BillToCode")
	If Request("ShipToCode") <> "" Then cmd("@ShipToCode") = Request("ShipToCode")
	cmd("@UserType") = userType
	cmd("@OP") = "O"
	set rs = Server.CreateObject("ADODB.recordset")
	set rs = cmd.execute()
	'If rs("IsBillValid") = "N" or rs("IsShipValid") = "N" Then Response.Redirect "../cart.asp"
End If

set rs = Server.CreateObject("ADODB.recordset")

confirm = ""

If Session("RetVal") <> "" Then	RetVal = Session("RetVal") Else	RetVal = Session("ConfRetVal")

sql = 	"declare @ObjectCode int set @ObjectCode = (select Object from R3_ObsCommon..TLOG where LogNum = " & RetVal & ") " & _
		"if exists(select 'A' from OLKCIC where InvLogNum = " & RetVal & ") begin set @ObjectCode = 48 end " & _
		"select @ObjectCode ObjCode, (select ViewObjPrint"
		
If userType = "C" Then sql = sql & "Client"

sql = sql & " from OLKDocConf where ObjectCode = @ObjectCode) ViewObjPrint "
set rs = conn.execute(sql)
ObjCode = rs("ObjCode")
ViewObjPrint = rs("ViewObjPrint") = "Y"
If myAut.GetObjectProperty(rs("ObjCode"), "C") and userType = "V" or Request("Confirm") = "Y" Then
	confirm = "H"
Else
	confirm = "C"
End If 

If userType = "C" Then

	If Not Session("noLic") Then
		set oLic = server.CreateObject("TM.LicenceConnect.LicenceConnection")
		oLic.LicenceServer = licip
		oLic.LicencePort = licport
		strResp = oLic.HasTrans(50, 1)
	Else
		strResp = "YES"
	End If
	If strResp <> "YES" Then
		response.redirect "cart.asp&update=Y&LicErr=" & strResp & "&ViewMode=" & Request("ViewMode")
	End If
End If

Series = myAut.GetObjectProperty(ObjCode, "S")
Series2 = myAut.GetObjectProperty(48, "S2")
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKCartSetDocFinalData" & Session("ID")
cmd.Parameters.Refresh()
cmd("@LanID") = Session("LanID")
cmd("@LogNum") = RetVal
cmd("@UserType") = userType
cmd("@branch") = Session("branch")
cmd("@ObjectCode") = ObjCode
If Series <> "NULL" Then cmd("@Series") = Series
If Series2 <> "NULL" Then cmd("@Series2") = Series2
cmd("@SumDec") = myApp.SumDec
cmd("@SlpCode") = Session("vendid")
cmd.execute()

If confirm = "H" Then
	If Request("Flow") <> "Y" Then
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKCreateUAFControl" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@UserType") = userType
		cmd("@ExecAt") = "D3"
		cmd("@ObjectEntry") = Session("RetVal")
		cmd("@AgentID") = Session("vendid")
		cmd("@LanID") = Session("LanID")
		cmd("@branch") = Session("branch")
		cmd("@SetLogNumConf") = "Y"
		cmd.execute()
	End If
	
	Session("ConfRetVal") = Session("RetVal")
	Session("RetVal") = ""
	
	redirUrl = "cartConfirm.asp?status=H"
	
	If Session("PayCart") Then 
		redirUrl = redirUrl & "&payment=Y"
		Session("ConfPayRetVal") = Session("PayRetVal")
		Session("PayRetVal") = ""
	End If
	
	Response.Redirect redirUrl
End If

sql = "select Status, Draft from R3_ObsCommon..TLOG where LogNum = " & RetVal
set rs = conn.execute(sql)
If rs("Status") = "R" Then
	Session("NotifyAdd") = True
	
	sqlclose = "declare @LogNum int set @LogNum = " & RetVal & " " & _
				"update R3_ObsCommon..TLOG set Status = 'C' where LogNum = @LogNum "
	

	'If userType = "C" Then
	'	sqlclose = sqlclose & "declare @mailID int set @mailID = IsNull((select Max(mailID)+1 from OLKMail), 0) " & _
	'							"insert OLKMail(mailID, TypeID, LanID, Entry, Sent) values(@mailID, 6, " & Session("LanID") & ", @LogNum, 'N') "
	'End If

	conn.execute(sqlclose)
End If	

txtConfAddDoc = getcartSubmitLngStr("LtxtConfAddDoc")
txtCreateNewDoc = getcartSubmitLngStr("LtxtCreateNewDoc")

Select Case ObjCode
	Case 13
		objDesc = txtInv
	Case 15
		objDesc = txtOdln
	Case 17
		objDesc = txtOrdr
		txtConfAddDoc = getcartSubmitLngStr("LtxtConfAddDoc2")
		txtCreateNewDoc = getcartSubmitLngStr("LtxtCreateNewDoc2")
	Case 23
		objDesc = txtQuote
	Case 48
		objDesc = txtInv & "/" & txtRct
	Case 203
		objDesc = txtODPIReq
	Case 204
		objDesc = txtODPIInv
	Case 22
		objDesc = txtOpor
End Select

docDesc = objDesc
isDraft = False
If rs("Draft") = "Y" Then
	docDesc = docDesc & " (" & getcartSubmitLngStr("DtxtDraft") & ")"
	isDraft = True
End If

set mySubmit = new SubmitControl
mySubmit.EnableRunInBackground = Request("I") <> "I3"
mySubmit.LogNum = RetVal 
mySubmit.LogNumID = "RetVal"
If Request("I") = "I3" Then
	If Session("PayRetVal") <> "" Then PayRetVal = Session("PayRetVal") Else PayRetVal = Session("ConfPayRetVal")
	mySubmit.LogNum2 = PayRetVal
	mySubmit.LogNumID2 = "PayRetVal"
End If
mySubmit.TransactionOkMessage = Replace(txtConfAddDoc, "{1}", docDesc)
Select Case userType
	Case "C"
		mySubmit.EndButtonDescription = getcartSubmitLngStr("LtxtContShop")
	Case "V"
		mySubmit.EndButtonDescription = Replace(txtCreateNewDoc, "{0}", objDesc)
End Select
If ViewObjPrint Then 
	mySubmit.SecondButtonDescription = getcartSubmitLngStr("LtxtView") & " " & objDesc
	mySubmit.SecondButtonFunction = "viewDetails({0});"
End If
If Request("I") = "I3" Then 
	mySubmit.RunInBackgroundRedir = "cartConfirm.asp?status=C&status=C&payment=Y"
	mySubmit.EndButtonFunction = "window.location.href='ventas/newCashInv.asp?AddPath=../';"
Else
	mySubmit.RunInBackgroundRedir = "cartConfirm.asp?status=C&status=C"
	Select Case userType
		Case "C"
			mySubmit.EndButtonFunction = "document.formSmallSearch.B2.click();"
		Case "V"
			mySubmit.EndButtonFunction = "window.location.href='ventas/newDocGo.asp?obj=" & ObjCode & "&AddPath=../';"
	End Select
End If
mySubmit.GenerateSubmit
isEntry = "N"
If Not isDraft Then
	If ObjCode = 48 Then 
		ObjCode = 13
	End If
Else
	ObjCode = 112
	isEntry = "Y"
End If %>
<script type="text/javascript">
function Start(page) 
{
	OpenWin = this.open(page, 'objDetails', 'toolbar=no,menubar=no,location=no,scrollbars=yes,resizable=yes, height=600,width=800');
	OpenWin.focus()
}

function viewDetails(objEntry)
{
	Start('');
	doMyLink('cxcDocDetailOpen.asp', 'DocType=<%=ObjCode%>&DocEntry=' + objEntry + '&pop=Y&AddPath=../&isEntry=<%=isEntry%><% If Request("I") = "I3" Then %>&payment=Y<% End If %>', 'objDetails');
}
</script>
<% Select Case userType
Case "C" %><!--#include file="clientBottom.asp"-->
<% Case "V" %><!--#include file="agentBottom.asp"-->
<% End Select %>
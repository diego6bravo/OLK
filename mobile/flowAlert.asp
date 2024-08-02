<!--#include file="lang/flowAlert.asp" -->
<%
sql = 	"select FlowID, Name, Type, " & _
		"ExecAt, Case When LineQuery is not null Then 'Y' Else 'N' End LineQry, NoteBuilder, NoteText, Draft, Authorize " & _
		"from OLKUAF " & _
		"where FlowID in (" & Request("DocFlowErr") & ") "

		If Request("DocConf") <> "" Then sql = sql & " and FlowID not in (" & Request("DocConf") & ") "

		sql = 	sql & "order by Type asc, [Order] asc"
rs.open sql, conn, 3, 1 %>
<style>
.GeneralTlt{
	font-family: Verdana;
	font-size: 11px;
	font-weight: bold;
	color:#4A4A4A;
	background-color: #A2D1FD;
	}
.GeneralTblBold{
	font-family: Verdana;
	font-size: 10px;
	font-weight:bold;
	color:#4A4A4A;
	background-color: #EDF5FE;
	}
</style>
<% If Request("myType") <> "2" Then %>
<table border="0" cellpadding="0" width="100%">
	<form method="POST" name="frmConf" action="operaciones.asp">
	<% 
	Draft = "N"
	Authorize = "N"
	do while not rs.eof
	FlowID = rs("FlowID")
	ExecAt = rs("ExecAt")
	If rs("Draft") = "Y" Then Draft = "Y"
	If rs("Authorize") = "Y" Then Authorize = "Y" %>
	<tr class="GeneralTlt">
		<td colspan="2">
		<%=rs("Name")%>&nbsp;</td>
	</tr>
	<tr class="GeneralTbl">
		<td width="83%" valign="top"><%=BuildNote(rs("ExecAt"))%>&nbsp;</td>
		<td width="17%" valign="top">
		<p align="center">
		<% Select Case rs("Type")
			Case 1 %>
		<img border="0" src="images/questionicon.gif" width="68" height="65" alt="<%=getflowAlertLngStr("LaltConf")%>">
		<%	Case 0 %>
		<img border="0" src="images/erroricon.gif" width="68" height="65" alt="<%=getflowAlertLngStr("LaltError")%>">
		<% 	Case 2 %>
		<img border="0" src="images/confirmicon.gif" width="68" height="65" alt="<%=getflowAlertLngStr("LaltFlow")%>">
		<% End Select %>
		</td>
	</tr>
	<% If rs("Type") = 2 Then %>
	<tr class="GeneralTblBold">
		<td colspan="2">
		<%=getflowAlertLngStr("DtxtNote")%>&nbsp;<%=rs.bookmark%>: 
		<input class="input" type="text" maxlength="254" size="94" name="ConfirmNote<%=rs("FlowID")%>" style="width: 100%"></td>
	</tr>
	<% End If %>
	<tr class="GeneralTbl">
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr class="GeneralTbl">
		<td colspan="2">
		&nbsp;</td>
	</tr>
	<tr class="GeneralTbl">
		<td colspan="2">&nbsp;</td>
	</tr>
	<% If rs("LineQry") = "Y" Then %>
	<tr class="GeneralTbl">
		<td colspan="2">
		<p align="center">
		<iframe name="content" width="100%" src='flowAlertDetails.asp?FlowID=<%=rs("FlowID")%>&amp;cmd=<%=Request("cmd")%>&amp;ExecAt=<%=rs("ExecAt")%>&amp;Item=<%=Request("FlowItem")%>&amp;SelDes=<%=SelDes%>&amp;WhsCode=<%=Request("FlowWhsCode")%>&amp;SaleType=<%=Request("FlowSaleType")%>&Quantity=<%=Request("FlowQuantity")%>&Price=<%=Request("Flowprecio")%>&chkAddAll=<%=Request("FlowSellAll")%>' border="0" frameborder="0" height="103">
		Your browser does not support inline frames or is currently configured not to display inline frames.
		</iframe></td>
	</tr>
	<% End If
	Select Case rs("Type")
		Case 0
			btnGoDesc = ""
			btnCancel = getflowAlertLngStr("LtxtReturn")
			myType = 0
		Case 1
			btnGoDesc = getflowAlertLngStr("DtxtYes")
			btnCancel = getflowAlertLngStr("DtxtNo")
			myType = 1
		Case 2
			btnGoDesc = getflowAlertLngStr("DtxtConfirm")
			btnCancel = getflowAlertLngStr("DtxtCancel")
			myType = 2
	End Select
	If rs("Type") <> 2 Then Exit Do
	rs.movenext
	loop %>
	<tr class="GeneralTbl">
		<td colspan="2">
		<p align="center"><% If btnGoDesc <> "" Then %>
		<input type="button" value="<%=btnGoDesc%>" name="B1" onclick="javascript:Confirm(<%=FlowID%>);" style="width: 80"> 
		-<% End If %>
		<input type="button" value="<%=btnCancel%>" name="B2" onclick="javascript:history.go(-1);" style="width: 80"></td>
	</tr>
	<input type="hidden" name="DocFlowErr" value="<%=Request("DocFlowErr")%>">
	<input type="hidden" name="DocConf" value="<%=Request("DocConf")%>">
	<input type="hidden" name="myType" value="<%=myType%>">
	<input type="hidden" name="retURL" value="<%=Request("retURL")%>">
	<input type="hidden" name="cmd" value="DocFlowErr">
	</form>
</table>
<% Select Case ExecAt
	Case "D1"
		frmAction = "newdocgonow.asp"
	Case "D2"
		If Request("ItemFastAdd") = "Y" Then
			frmAction = "cart/cartFastAddSubmit.asp"
		Else
			frmAction = "cart/addcartsubmit.asp"
		End If
	Case "D3"
		frmAction = "cart/cartupdate.asp"
	Case "C1"
		frmAction = "client/submitClient.asp"
	Case "C2"
		frmAction = "activity/actSubmit.asp"
End Select %>
<form action="<%=frmAction%>" method="post" name="frmAddCart">
<% For each itm in Request.QueryString
If itm <> "DocFlowErr" and itm <> "retURL" and itm <> "DocConf" and itm <> "ItemFastAdd" Then %>
<input type="hidden" name="<%=itm%>" value="<%=myHTMLEncode(Request(itm))%>">
<% End If
Next %>
<% For each itm in Request.Form
If itm <> "DocFlowErr" and itm <> "retURL" and itm <> "DocConf" and itm <> "ItemFastAdd" Then %>
<input type="hidden" name="<%=itm%>" value="<%=myHTMLEncode(Request(itm))%>">
<% End If
Next %>
<% If Request("retURL") <> "" Then
	ArrVal = Split(Request("retURL"), "{a}")
	For i = 0 to UBound(ArrVal)
	ArrVal2 = Split(ArrVal(i), "{e}")
	VarName = ArrVal2(0)
	VarVal = ArrVal2(1)
	If VarName <> "confirm" or VarName = "confirm" and myType <> 2 Then %>
	<input type="hidden" name="<%=ArrVal2(0)%>" value="<%=myHTMLEncode(ArrVal2(1))%>">
	<% 
	End If
	Next %>
<input type="hidden" name="isFlow" value="Y">
<input type="hidden" name="DocConf" value="">
<input type="hidden" name="Draft" value="<%=Draft%>">
<input type="hidden" name="Authorize" value="<%=Authorize%>">
<% If myType = 2 Then %><input type="hidden" name="Confirm" value="Y"><% End If %>
</form>
<% End If %>
<script language="javascript">
function Confirm(FlowID)
{
	if (document.frmConf.DocConf.value != '') { document.frmConf.DocConf.value += ', '; }
	document.frmConf.DocConf.value += FlowID;
	<% 
	If myType = 1 Then %>
		<% If rs.recordcount > 1 Then %>
			document.frmConf.submit();
		<% Else %>
			document.frmAddCart.DocConf.value = document.frmConf.DocConf.value;
			document.frmAddCart.submit();
		<% End If %>
	<% ElseIf myType = 2 Then %>
			document.frmAddCart.DocConf.value = document.frmConf.DocConf.value;
			document.frmAddCart.submit();
	<% End If %>
}
</script>
<% Else
	rs.Filter = "Type = 2"
	If Request("cmd") = "cart" Then 
		LogNum = Session("RetVal")
	ElseIf Request("cmd") = "payment" Then
		LogNum = Session("PayRetVal")
	ElseIf Request("cmd") = "newItem" Then
		LogNum = Session("ItmRetVal")
	ElseIf Request("cmd") = "newClient" Then
		LogNum = Session("CrdRetVal")
	End If
	sql = "declare @LogNum int set @LogNum = " & LogNum & " declare @FlowID int declare @Note nvarchar(254) "
	do while not rs.eof
		sql = sql & "set @FlowID = " & rs("FlowID") & " " & _
		"set @Note = '" & Left(Request("ConfirmNote" & rs("FlowID")),254) & "' " & _
		"if not exists(select 'A' from OLKUAF3 where LogNum = @LogNum and FlowID = @FlowID) and @Note is not null begin " & _
		"	insert OLKUAF3(LogNum, FlowID, Note) values(@LogNum, @FlowID, @Note) " & _
		"end else begin " & _
		"	update OLKUAF3 set Note = @Note where LogNum = @LogNum and FlowID = @FlowID " & _
		"end "
	rs.movenext
	loop
	conn.execute(sql) %>
<script language="javascript">
opener.goAdd('Y', '<%=Request("DocConf")%>');
window.close();
</script>
<% End If %>
<% Function BuildNote(ByVal ExecAt)
	myNote = rs("NoteText")
	If rs("NoteBuilder") = "Y" Then
	set cmd = Server.CreateObject("ADODB.Command")
		set rNote = Server.CreateObject("ADODB.RecordSet")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKCheckDF" & Session("ID") & "_" & Replace(rs("FlowID"), "-", "_") & "_msg"
		LoadCmdParams
		set rNote = cmd.execute()
		If Not rNote.Eof Then
			For each fld in rNote.Fields
				If Not IsNull(fld) Then myNote = Replace(myNote, "{" & fld.Name & "}", fld) Else myNote = Replace(myNote, "{" & fld.Name & "}", "")
			Next
		End If
	End If

	myNote = Replace(myNote, chr(13), "<br>")
	BuildNote = myNote
End Function

Sub LoadCmdParams
	cmd.Parameters.Refresh()

	cmd("@LanID") = Session("LanID")
	cmd("@SlpCode") = Session("vendid")
	cmd("@dbName") = Session("olkdb")
	cmd("@branch") = Session("branch")
	cmd("@UserType") = "V"
	
	Select Case ExecAt
		'Case "O0", "O1", "O7" ' Aprove Sales Order, Convert Quotation to Sales Order, Convert Sales Order to Invoice
		'	cmd("@Entry") = arrVars(0)
		'Case "O2", "O3", "O4" ' Close  Object, Cancel Object, Remove Object
		'	cmd("@ObjectCode") = arrVars(0) 
		'	cmd("@Entry") = arrVars(1)
		Case "D2" ' Add Item
			cmd("@LogNum") = Session("RetVal")
			cmd("@CardCode") = Session("UserName")
			cmd("@ItemCode") = Request("FlowItem")
			If Request("FlowQuantity") <> "" Then cmd("@Quantity") = CDbl(getNumericOut(Request("FlowQuantity")))
			If Request("FlowSaleType") <> "" Then cmd("@Unit") = Request("FlowSaleType")
			If Request("Flowprecio") <> "" Then cmd("@Price") = CDbl(getNumericOut(Request("Flowprecio")))
			If Request("FlowWhsCode") <> "" Then cmd("@WhsCode") = Request("FlowWhsCode")
			If Request("FlowSellAll") = "Y" Then cmd("@All") = "Y"
		Case "D3" ' LtxtDocConf
			cmd("@LogNum") = Session("RetVal")
		Case "R1" ' LtxtCreation	******clean*******
		Case "R2" ' LtxtRcpConf
			cmd("@LogNum") = Session("PayRetVal")
		Case "A1" ' LtxtItmConf
			cmd("@LogNum") = Session("ItmRetVal")
		Case "C1" ' LtxtClientConf
			cmd("@LogNum") = Session("CrdRetVal")
		Case "C2" ' LtxtActivityConf
			cmd("@LogNum") = Session("ActRetVal")
		Case "C3" ' LtxtActivityConf
			cmd("@LogNum") = Session("SORetVal")
	End Select	
	
	Select Case ExecAt
		Case "C2", "C3", "R1", "R2", "D3", "D1"
			cmd("@CardCode") = Session("UserName")
	End Select
End Sub
 %>
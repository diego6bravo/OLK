<!--#include file="chkLogin.asp" -->
<!--#include file="lang/flowAlert.asp" -->
<!--#include file="myHTMLEncode.asp"-->
<!--#include file="lcidReturn.inc"-->
<html <% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>>
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

set rs = Server.CreateObject("ADODB.RecordSet")
If userType = "V" Then
	SelDes = "0"
ElseIf userType = "C" Then
	sql = "select SelDes from OLKCommon"
	set rs = conn.execute(sql)
	SelDes = rs(0)
	rs.close
End If

Select Case Request("cmd")
	Case "cart" 
		LogNum = Session("RetVal")
	Case "payment" 
		LogNum = Session("PayRetVal")
	Case "newItem" 
		LogNum = Session("ItmRetVal")
	Case "newClient" 
		LogNum = Session("CrdRetVal")
	Case "newActivity" 
		LogNum = Session("ActRetVal")
End Select

If Request("Draft") = "Y" Then
	sql = "update R3_ObsCommon..TLOG set Draft = 'Y' where LogNum = " & LogNum 
	conn.execute(sql)
End If
If Request("Authorize") = "Y" Then

End If

sql = 	"select FlowID, Name, Case When '" & userType & "' = 'C' and Type = 1 and ExecAt = 'C1' Then 0 Else Type End Type, " & _
		"ExecAt, LineQuery, NoteBuilder, NoteQuery, NoteText, Draft, Authorize " & _
		"from OLKUAF " & _
		"where FlowID in (" & Request("DocFlowErr") & ") "

		If Request("DocConf") <> "" Then sql = sql & " and FlowID not in (" & Request("DocConf") & ") "

		sql = 	sql & "order by Type asc, [Order] asc"
rs.open sql, conn, 3, 1 %>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" type="text/css" href="design/<%=SelDes%>/style/stylePopUp.css">
<title><%=getflowAlertLngStr("LttlDocFlow")%></title>
<script language="javascript">var doNoLang = true;</script>
<script language="javascript" src="ventas.js"></script>
</head>

<body onload="window.resizeTo(550,400)">
<% If Request("myType") <> "2" Then %>
<table border="0" cellpadding="0" width="100%" id="table1">
	<form method="POST" name="frmConf" action="flowAlert.asp">
	<% 
	Draft = "N"
	Authorize = "N"
	do while not rs.eof %>
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
		<img border="0" src="design/<%=SelDes%>/images/questionicon.gif" width="68" height="65" alt="<%=getflowAlertLngStr("LaltConf")%>">
		<%	Case 0 %>
		<img border="0" src="design/<%=SelDes%>/images/erroricon.gif" width="68" height="65" alt="<%=getflowAlertLngStr("LaltError")%>">
		<% 	Case 2 %>
		<img border="0" src="design/<%=SelDes%>/images/confirmicon.gif" width="68" height="65" alt="<%=getflowAlertLngStr("LaltFlow")%>">
		<% End Select %>
		</td>
	</tr>
	<% If rs("Type") = 2 Then %>
	<tr class="GeneralTblBold">
		<td colspan="2">
		<%=getflowAlertLngStr("DtxtNote")%>&nbsp;<%=rs.bookmark%>: 
		<input class="input" type="text" size="94" name="ConfirmNote<%=rs("FlowID")%>" onkeydown="return chkMax(event, this, 256);"></td>
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
	<% If Not IsNull(rs("LineQuery")) Then %>
	<tr class="GeneralTbl">
		<td width="777" colspan="2">
		<p align="center">
		<iframe name="content" width="100%" src="flowAlertDetails.asp?FlowID=<%=rs("FlowID")%>&cmd=<%=Request("cmd")%>&ExecAt=<%=rs("ExecAt")%>&Item=<%=Server.URLEncode(Request("Item"))%>&SelDes=<%=SelDes%>&WhsCode=<%=Request("WhsCode")%>&SaleType=<%=Request("SaleType")%>&addQty=<%=Request("addQty")%>&precio=<%=Request("precio")%>" border="0" frameborder="0" height="103">
		Your browser does not support inline frames or is currently configured not to display inline frames.
		</iframe></td>
	</tr>
	<% End If
	Select Case rs("Type")
		Case 0
			btnGoDesc = ""
			btnCancel = getflowAlertLngStr("DtxtClose")
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
	FlowID = rs("FlowID")
	ExecAt = rs("ExecAt")
	If rs("Draft") = "Y" Then Draft = "Y"
	If rs("Authorize") = "Y" Then Authorize = "Y"
	If rs("Type") <> 2 Then Exit Do
	rs.movenext
	loop %>
	<tr class="GeneralTbl">
		<td width="777" colspan="2">
		<p align="center"><% If btnGoDesc <> "" Then %>
		<input type="button" value="<%=btnGoDesc%>" name="B1" onclick="javascript:Confirm(<%=FlowID%>);" style="width: 60"> 
		-<% End If %>
		<input type="button" value="<%=btnCancel%>" name="B2" onclick="javascript:window.close()" style="width: 60"></td>
	</tr>
	<input type="hidden" name="DocFlowErr" value="<%=Request("DocFlowErr")%>">
	<input type="hidden" name="DocConf" value="<%=Request("DocConf")%>">
	<input type="hidden" name="myType" value="<%=myType%>">
	<input type="hidden" name="cmd" value="<%=Request("cmd")%>">
	<input type="hidden" name="Draft" value="<%=Draft%>">
	<input type="hidden" name="Authorize" value="<%=Authorize%>">
	</form>
</table>
<% If ExecAt = "D2" Then %>
<form action="cart/addCartSubmit<% If Request("retURL") <> "" Then %>M<% End If %>.asp" method="post" name="frmAddCart">
<% For each itm in Request.QueryString
If itm <> "DocFlowErr" and itm <> "retURL" Then %>
<input type="hidden" name="<%=itm%>" value="<%=Server.HTMLEncode(Request(itm))%>">
<% End If
Next %>
<% For each itm in Request.Form
If itm <> "DocFlowErr" and itm <> "retURL" Then %>
<input type="hidden" name="<%=itm%>" value="<%=Server.HTMLEncode(Request(itm))%>">
<% End If
Next %>
<% If Request("retURL") <> "" Then
	ArrVal = Split(Request("retURL"), "{a}")
	For i = 0 to UBound(ArrVal)
	ArrVal2 = Split(ArrVal(i), "{e}") %>
<input type="hidden" name="<%=ArrVal2(0)%>" value="<%=myHTMLEncode(ArrVal2(1))%>">
<% Next
End If %>
<input type="hidden" name="isFlow" value="Y">
<input type="hidden" name="DocConf" value="">
</form>
<% End If %>
<script language="javascript">
function Confirm(FlowID)
{
	<% If myType <> 2 Then %>
		if (document.frmConf.DocConf.value != '') { document.frmConf.DocConf.value += ', '; }
		document.frmConf.DocConf.value += FlowID;
	<% End If
	If myType = 1 Then %>
		<% If rs.recordcount > 1 Then %>
			document.frmConf.submit();
		<% Else %>
			<% If ExecAt = "D2" Then %>
			document.frmAddCart.DocConf.value = document.frmConf.DocConf.value;
			document.frmAddCart.submit();
			<% Else %>
			opener.goAdd('N', document.frmConf.DocConf.value<% If Request("cmd") = "cart" or Request("cmd") = "payment" or Request("cmd") = "cartAddresses" Then %>, '<%=Draft%>', '<%=Authorize%>'<% End If %>);
			window.close();
			<% End If %>
		<% End If %>
	<% ElseIf myType = 2 Then %>
		document.frmConf.submit();
	<% End If %>
}
</script>
<% Else
	rs.Filter = "Type = 2"
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
opener.goAdd('Y', '<%=Request("DocConf")%>'<% If Request("cmd") = "cart" Then %>, '<%=Draft%>', '<%=Authorize%>'<% End If %>);
window.close();
</script>
<% End If %>
</body>

</html>
<% Function BuildNote(ByVal ExecAt)
	myNote = rs("NoteText")
	If rs("NoteBuilder") = "Y" Then
		If Left(ExecAt,1) = "D" Then 
			LogNum = Session("RetVal")
		ElseIf Left(ExecAt,1) = "R" Then
			LogNum = Session("PayRetVal")
		ElseIf ExecAt = "C2" Then
			LogNum = Session("ActRetVal")
		End If
		
		If ExecAt <> "D1" and ExecAt <> "R1" Then sqlBase = "declare @LogNum int set @LogNum = " & LogNum & " "
		
		sqlBase = 	sqlBase & "declare @SlpCode int set @SlpCode = " & Session("VendID") & " " & _
					"declare @dbName nvarchar(100) set @dbName = db_name() " & _
					"declare @branch int set @branch = " & Session("branch") & " " & _
					"declare @LanID int set @LanID = " & Session("LanID") & " "
					
		If Left(ExecAt,1) = "D" or Left(ExecAt,1) = "R" or ExecAt = "C2" Then sqlBase = sqlBase & "declare @CardCode nvarchar(15) set @CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "' "
		
		If ExecAt = "D2" Then 
			If Request("SaleType") <> "" Then SaleType = Request("SaleType") Else SaleType = "NULL"
			If Request("WhsCode") <> "" Then WhsCode = "N'" & Request("WhsCode") & "'" Else WhsCode = "OLKCommon.dbo.DBOLKGetWhsCode" & Session("ID") & "(" & Session("branch") & ", " & Session("vendid") & ", @ItemCode)"
			If Request("precio") <> "" Then precio = getNumeric(Request("precio")) Else precio = "NULL"
			If Request("chkAddAll") <> "Y" Then 
				If Request("addQty") <> "" Then addQty = getNumeric(Request("addQty")) Else addQty = "1"
			Else
				addQty = "OLKCommon.dbo.DBOLKItemInv" & Session("ID") & "Val(@ItemCode, @WhsCode, @dbName, @LogNum, -1)"
			End If
			
			sqlBase = sqlBase & "declare @ItemCode nvarchar(20) set @ItemCode = N'" & Request("Item") & "' " & _
								"declare @WhsCode nvarchar(8) set @WhsCode = " & WhsCode & " " & _
								"declare @Unit smallint set @Unit = " & SaleType & " " & _
								"If @Unit is null Begin " & _
								"set @Unit = Case '" & userType & "'  " & _
								"		When 'C' Then (select ClientSaleUnit from olkcommon)  " & _
								"		When 'V' Then (select AgentSaleUnit from olkcommon) End End " & _
								"declare @Quantity numeric(19,6) set @Quantity = " & addQty & " " & _
								"declare @Price numeric(19,6) set @Price = " & precio & " " & _
								"If @Price is null begin " & _
								"	EXEC OLKCommon..DBOLKGetItemPrice" & Session("ID") & " @ItemCode = @ItemCode, @CardCode = @CardCode, @PriceList = " & Session("PriceList") & ", @UserType = '" & userType & "', @ItemPrice = @Price out " & _
								"End If @Price is null begin set @Price = 0 End "					
		End If
					
		sql = sqlBase & rs("NoteQuery")
		sql = QueryFunctions(sql)
		set rNote = Server.CreateObject("ADODB.RecordSet")
		set rNote = conn.execute(sql)
		If Not rNote.Eof Then
			For each item in rNote.Fields
				If Not IsNull(item) Then
					myNote = Replace(myNote,"{" & item.Name & "}", item)
				End If
			next
		End IF
	End If
	set rNote = nothing
	myNote = Replace(myNote, chr(13), "<br>")
	BuildNote = myNote
End Function %>
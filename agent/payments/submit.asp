<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="../chkLogin.asp" -->
<!--#include file="../lcidReturn.inc"-->
<!--#include file="../authorizationClass.asp"-->
<html>

<body>
<!--#include file="../linkForm.asp"-->
<%

Dim myAut
set myAut = New clsAuthorization

Dim close
varx = 0

set rs = Server.CreateObject("ADODB.recordset")
           
Select Case Request("submitCmd") 
	Case "update" 
		If Request("Draft") = "Y" Then
			sql = "update R3_ObsCommon..TLOG set Draft = 'Y' where LogNum = " & Session("PayRetVal")
			conn.execute(sql)
		End If
		
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKPostObjectCreation" & Session("ID")
		cmd.Parameters.Refresh
		cmd("@sessiontype") = "A"
		cmd("@object") = 24
		cmd("@LogNum") = Session("PayRetVal")
		cmd("@transtype") = "A"
		cmd("@CurrentSlpCode") = Session("vendid")
		cmd("@Branch") = Session("branch")
		cmd.execute()
	
		sql = 	"select IsNULL(CashSum,0)+ IsNULL(TrsfrSum,0)+ " & _
				"IsNULL((select sum(CreditSum) from r3_obscommon..pmt3 where Lognum = T0.Lognum),0)+ " & _
				"IsNULL((select sum(CheckSum) from r3_obscommon..pmt1 where lognum = T0.LogNum),0) " & _
				"from r3_obscommon..tpmt T0 where Lognum = " & Session("PayRetVal")
		set rs = conn.execute(sql)
		
		If CDbl(rs(0)) > 0 Then
			If Request("Confirm") = "N" Then
				If myAut.GetObjectProperty(24, "C") Then 
					sqlAdd1 = "'Y'" 
					goStatus = "H"
				Else 
					sqlAdd1 = "'N'"
					goStatus = "C"
				End If
			Else
				sqlAdd1 = "'Y'"
				goStatus = "H"
			End If
			
			sql = 	"declare @LogNum int set @LogNum = " & Session("PayRetVal") & " " & _
					"If (select ORCTContraComp from olkcommon) = 'Y' Begin " & _
					"	update r3_obscommon..tpmt set CounterRef = LogNum where LogNum = @LogNum " & _
					"End " & _
					"update R3_ObsCommon..TLOGControl set SlpCode = " & Session("vendid") & ", ConfBranch = " & Session("branch") & ", Background = 'N' where LogNum = @LogNum "
			conn.execute(sql)

			If goStatus = "Y" Then
				set cmd = Server.CreateObject("ADODB.Command")
				cmd.ActiveConnection = connCommon
				cmd.CommandType = &H0004
				cmd.CommandText = "DBOLKCreateUAFControl" & Session("ID")
				cmd.Parameters.Refresh()
				cmd("@UserType") = "V"
				cmd("@ExecAt") = "R2" 
				cmd("@ObjectEntry") = Session("PayRetVal")
				cmd("@AgentID") = Session("vendid")
				cmd("@LanID") = Session("LanID")
				cmd("@branch") = Session("branch")
				cmd("@SetLogNumConf") = "Y"
				cmd.execute()
			Else
				set cmd = Server.CreateObject("ADODB.Command")
				cmd.ActiveConnection = connCommon
				cmd.CommandType = &H0004
				cmd.CommandText = "DBOLKStartLogProcess"
				cmd.Parameters.Refresh()
				cmd("@LogNum") = Session("PayRetVal")
				cmd("@ErrCode") = GetLangErrCode()
				cmd.execute()
			End If
			
			Select Case goStatus 
				Case "C" 
					Session("NotifyAdd") = True
					response.redirect "../agentPaymentSubmit.asp?Confirm=" & Request("Confirm")
				Case "H"
					Session("ConfPayRetVal") = Session("PayRetVal")
					Session("PayRetVal") = ""
					response.redirect "../agentPaymentConfirm.asp?Confirm=" & goStatus
			End Select
		Else
			response.redirect "../agentPayment.asp?err=importe"
		End If
	Case "canceldoc" 
			response.redirect "../ventas/docDel.asp?retval=" & Session("PayRetVal")
	Case "payCash" 
			If Request("Total") <> "" Then 
				CashSum = ClearFormatNumber(Request("Total"), Request("DocCur"))
			Else 
				CashSum = "NULL"
			End If
			sql = "update r3_obscommon..tpmt set CashAcct = N'" & Request("cuenta") & "', cashsum = " & CashSum & " where lognum = " & Session("PayRetVal")
			conn.execute(sql) 
			If Request("cashInv") = "Y" Then
				sql = "update olkcic set cashsum = " & getNumeric(Request("CashTotal")) & " where PayLogNum = " & Session("PayRetVal")
				conn.execute(sql)
			End If
			updatePagado ""
	Case "payTrans" 
		If Request("Total") <> "" and Request("Total") <> "0" Then
		   	TransferSum = ClearFormatNumber(Request("Total"), Request("DocCur"))

			sql = "update r3_obscommon..tpmt set TrsfrAcct = N'" & Request("AcctCode") & _
				  "', TrsfrSum = " & TransferSum & ", TrsfrDate = Convert(datetime,'" & SaveSqlDate(Request("ftrans")) & _
				  "',120), TrsfrRef = N'" & Request("comp") & "' where lognum = " & Session("PayRetVal")
			conn.execute(sql)
			updatePagado ""
		Else
			sql = "update r3_obscommon..tpmt set TrsfrAcct = NULL, TrsfrSum = NULL, " & _
				  "TrsfrDate = NULL, TrsfrRef = NULL"
			conn.execute(sql)
			updatePagado ""
		End If
	Case "addCred" 
		If Request("cpagoadd") <> "" Then cpagoadd = ClearFormatNumber(Request("cpagoadd"),Request("DocCur")) Else cpagoadd = "NULL"
		CardValidDate = "Cast('" & Request("CardValidM") & "/01/" & Request("CardValidY") & "' as datetime)"
		
		sql = 	"declare @linenum int set @linenum = ISNULL((select max(linenum) + 1 from r3_obscommon..pmt3 where lognum = " & Session("PayRetVal") & "),0) " & _
				"insert r3_obscommon..pmt3(LogNum, LineNum, CreditCard, CreditAcct, CrCardNum, CardValid, VoucherNum, OwnerIdNum, " & _
				"OwnerPhone, CrTypeCode, NumOfPmnts, FirstSum, AddPmntSum, CreditSum, ConfNum) " & _
				"Values(" & Session("PayRetVal") & ", @LineNum, N'" & Request("CreditCard") & "', N'" & Request("CreditAcct") & "', N'" & Request("CrCardNum") & "', " & _
				CardValidDate & ", N'" & Request("compnum") & "', N'" & Request("OwnerIdNum") & "', N'" & Request("OwnerPhone") & "', " & _
				"N'" & Request("SistPagCode") & "', N'" & Request("pagcant") & "', " & ClearFormatNumber(Request("perpago"),Request("DocCur")) & ", " & cpagoadd & ", " & _
				ClearFormatNumber(Request("impval"),Request("DocCur")) & ", N'" & Request("autorizacion") & "')"
		conn.execute(sql)
		updatePagado "payCred.asp?voucher=NULL&imp=" & Request("imp") & "&saldofuera=" & Request("saldofuera")
	Case "updateCred" 
		If Request("cpagoadd") <> "" Then cpagoadd = ClearFormatNumber(Request("cpagoadd"),Request("DocCur")) Else cpagoadd = "NULL"
		CardValidDate = "Cast('" & Request("CardValidM") & "/01/" & Request("CardValidY") & "' as datetime)-day(Cast('01/" & Request("CardValidM") & "/" & Request("CardValidY") & "' as datetime))"
		sql = 	"update r3_obscommon..pmt3 set CreditCard = N'" & Request("CreditCard") & "', CreditAcct = N'" & Request("CreditAcct") & "', CrCardNum = N'" & Request("CrCardNum") & "'," & _
				"CardValid = " & CardValidDate & ", VoucherNum = N'" & Request("compnum") & "', OwnerIdNum= N'" & Request("OwnerIdNum") & "', OwnerPhone = N'" & Request("OwnerPhone") & "', CrTypeCode = N'" & Request("SistPagCode") & "'," & _
				"NumOfPmnts = '" & Request("pagcant") & "', FirstSum = " & ClearFormatNumber(Request("perpago"),Request("DocCur")) & ", AddPmntSum = " & cpagoadd & ", CreditSum = " & ClearFormatNumber(Request("impval"),Request("DocCur")) & ", " & _
				"ConfNum = N'" & Request("autorizacion") & "' where lognum = " & Session("PayRetVal") & " and linenum = " & Request("linenum")
		conn.execute(sql)
		updatePagado "payCred.asp?voucher=NULL&imp=" & Request("imp") & "&saldofuera=" & Request("saldofuera")
	Case "payCred" 
		updatePagado ""
	Case "delCard" 
		sql = "delete r3_obscommon..pmt3 where lognum = " & Session("PayRetVal") & " and linenum = " & Request("linenum")
		conn.execute(sql)
		updatePagado "payCred.asp?voucher=NULL&imp=" & Request("imp") & "&saldofuera=" & Request("saldofuera")
	Case "payCheck" 
		If Request("delete.x") <> "" Then
			sql = "delete r3_obscommon..pmt1 where lognum = " & Session("PayRetVal") & " and LineNum = " & Request("linenum")
			conn.execute(sql)
			updatePagado "payCheck.asp?imp=" & Request("imp") & "&saldofuera=" & Request("saldofuera")
		Else
			sql = "update R3_ObsCommon..TPMT set CheckAcct = N'" & Request("CheckAcct") & "' where LogNum = " & Session("PayRetVal")
			conn.execute(sql)
			sql = "select LineNum from r3_obscommon..pmt1 where lognum = " & Session("PayRetVal")
			set rs = conn.execute(sql)
			If Not rs.Eof Then
				sql = ""
				do while not rs.eof
					If Request("fecha" & rs("LineNum")) <> "" Then fecha = "Convert(datetime,'" & SaveSqlDate(Request("fecha" & rs("LineNum"))) & "',120)" else fecha = "NULL"
					If Request("banco" & rs("LineNum")) <> "" Then banco = "N'" & Request("banco" & rs("LineNum")) & "'" else banco = "NULL"
					If Request("sucursal" & rs("LineNum")) <> "" Then sucursal = "N'" & Request("sucursal" & rs("LineNum")) & "'" else sucursal = "NULL"
					If Request("cuenta" & rs("LineNum")) <> "" Then cuenta = "N'" & Request("cuenta" & rs("LineNum")) & "'" else cuenta = "NULL"
					If Request("detalles" & rs("LineNum")) <> "" Then detalles = "N'" & Request("detalles" & rs("LineNum")) & "'" else detalles = "NULL"
		
				   	CheckSum = ClearFormatNumber(Request("imp" & rs("LineNum")), Request("DocCur"))
			   	
					sql = sql & "update r3_obscommon..pmt1 set DueDate = " & fecha & ", bankcode = " & banco & ", branch = " & sucursal & _
								", acctnum = " & cuenta & ", checknum = " & detalles & ", checksum = " & CheckSum & _
								" where lognum = " & Session("PayRetVal") & " and linenum = " & rs("linenum")
				rs.movenext
				loop
				conn.execute(sql)
			End If
			varReturn = ""
			If Request("impval") <> "" Then varReturn = addChk()
			updatePagado varReturn
		End If
End Select
Public Function addChk()
	If Request("impval") = "" Then
		err = 1
	ElseIf len(Request("cuenta")) > 50 Then
		err = 3
	ElseIf len(Request("detalles")) > 254 Then
		err = 4
	Else
		Err = 0
	End If
	
	If err = 0 Then
			If Request("fecha") <> "" Then fecha = "Convert(datetime,'" & SaveSqlDate(Request("fecha")) & "',120)" else fecha = "NULL"
			If Request("banco") <> "" Then banco = Request("banco") else banco = "NULL"
			If Request("sucursal") <> "" Then sucursal = "N'" & Request("sucursal") & "'" else sucursal = "NULL"
			If Request("cuenta") <> "" Then cuenta = "N'" & Request("cuenta") & "'" else cuenta = "NULL"
			If Request("detalles") <> "" Then detalles = "N'" & Request("detalles") & "'" else detalles = "NULL"
			sql = "declare @linenum int set @linenum = ISNULL((select max(linenum)+1 from r3_obscommon..pmt1 where lognum = " & Session("PayRetVal") & "),0) " & _
				  "insert r3_obscommon..pmt1(LogNum, LineNum, DueDate, BankCode, Branch, AcctNum, CheckNum, CheckSum) " & _
				  "Values(" & Session("PayRetVal") & ", @LineNum, " & fecha & ", '" & banco & "', " & sucursal & ", " & cuenta & _
				  ", " & detalles & ", " & Replace(getNumeric(Request("impval")),Request("DocCur"),"") & ")"
			conn.execute(sql)
			If Request("submitCmd") = "payCheck" and Request("Agregar") <> "" then 
				addChk = "payCheck.asp?imp=" & Request("imp") & "&saldofuera=" & Request("saldofuera")
			End If
	Else
			addChk = "payCheck.asp?err=" & err & "&fecha=" & Request("fecha") & "&banco=" & Request("banco") & _
							  "&sucursal=" & Request("sucursal") & "&cuenta=" & Request("cuenta") & "&detalles=" & Request("detalles") & _
							  "&impval=" & Request("impval") & "&imp=" & Request("imp")
	End If

End Function
set rs = nothing
conn.close 
Public Sub updatePagado(Link) 
sql =            "declare @lognum int set @lognum = " & Session("PayRetVal") & " " & _
           		 "declare @CardCode nvarchar(15) set @CardCode = (select CardCode from R3_ObsCommon..tpmt where lognum = @lognum) " & _
				 " select " & _
				 "(select ISNULL(sum(cashsum),0) from r3_obscommon..tpmt where lognum = @lognum)+ " & _
           		 "(select ISNULL(sum(checksum),0) from r3_obscommon..pmt1 where lognum = @lognum)+ " & _
           		 "(select ISNULL(sum(creditsum),0) from r3_obscommon..pmt3 where lognum = @lognum)+ " & _
           		 "(select ISNULL(sum(TrsfrSum),0) from r3_obscommon..tpmt where lognum = @lognum) Pagado, " & _
           		 "IsNULL(((select sum(Credit) from jdt1 where ShortName = @CardCode and TransType = 24)- " & _
				 "(select sum(PaidToDate) from oinv where CardCode = @CardCode)- " & _
				 "(select sum(Debit) from jdt1 where ShortName = @CardCode and TransType <> 13))*-1,0) SaldoFuera "
           		 set rs = conn.execute(sql)
           		 pagVal = CDbl(rs("Pagado"))
           		 If Request("saldofuera") = "true" Then pagVal = pagVal + (CDbl(rs("SaldoFuera"))*-1) %>
<script language="javascript" src="../general.js"></script>
<script language="javascript">
opener.updatePagado('<%=FormatNumber(pagVal,myApp.SumDec)%>');
<% If Link = "" Then %>
window.close();
<% ElseIf Link <> "" Then %>
//window.location.href='<%=Link%>';
doMyLink('<%=Split(Link, "?")(0)%>', '<%=Split(Link, "?")(1)%>', '');
<% end if %>
</script>
<% end sub %>

</body>
</html>
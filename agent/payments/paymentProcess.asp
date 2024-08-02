<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="../chkLogin.asp" -->
<!--#include file="../myHTMLEncode.asp"-->
<!--#include file="../lcidReturn.inc" -->
<%

LogNum = Session("PayRetVal")
Field = Request.Form("Field")
FieldType = Request.Form("FieldType")
Value = Request.Form("Value")

If Left(Field, 2) = "U_" Then
	set rs = Server.CreateObject("ADODB.RecordSet")
	sql = "select SDKID from R3_ObsCommon..TCIF where CompanyDB = db_name()"
	set rs = conn.execute(sql)
	Field = Replace(Field, "U_", rs("SDKID"))
End If

Select Case Field 
	Case "SaldoFuera" 
		Select Case Value
			Case "Y"
				sql = 	"if not exists(select 'A' from R3_ObsCommon..OLKTPMT where lognum = " & LogNum & ") Begin " & _
	        			"insert R3_ObsCommon..OLKTPMT (LogNum) values(" & LogNum & ") End "
			Case "N"
				sql = "delete R3_ObsCommon..OLKTPMT where lognum = " & LogNum
		End Select
		conn.execute(sql)
	Case "Check", "SumApplied"
		docType = Request.Form("DocType")
		docNum = Request.Form("DocNum")
		instID = Request.Form("InstID")
		
		sql = 	"declare @LogNum int set @LogNum = " & LogNum & " " & _
				"declare @DocType int set @DocType = " & docType & " " & _
				"declare @DocEntry int set @DocEntry = Case @DocType When 13 Then (select DocEntry from OINV where DocNum = " & docNum & ") " & _
				"													When 203 Then (select DocEntry from ODPI where DocNum = " & docNum & ") End " & _
				"declare @InstID int set @InstID = " & instID & " "
				
		Select Case Field
			Case "Check"
				Select Case Value
					Case "N"
						sql = sql & "delete R3_ObsCommon..PMT2 where LogNum = @LogNum and InvType = Convert(nvarchar(50),@DocType) and DocEntry = @DocEntry and InstId = @InstID "
				End Select
			Case "SumApplied"
				sql = sql & "declare @FC char(1) set @FC = Case When  " & _  
							"Case @DocType When 13 Then (select DocCur from OINV where DocEntry = @DocEntry) " & _  
							"When 203 Then (select DocCur from ODPI where DocEntry = @DocEntry) End = (select MainCurncy from OADM) Then 'N' Else 'Y' End "  & _
							"if not exists(select '' from R3_ObsCommon..PMT2 where LogNum = @LogNum and InvType = Convert(nvarchar(50),@DocType) and DocEntry = @DocEntry and InstId = @InstID) begin " & _
							"declare @LineNum int set @LineNum = IsNull((select Max(LineNum)+1 from R3_ObsCommon..PMT2 where LogNum = @LogNum), 0) " & _
							"insert R3_ObsCommon..PMT2(LogNum, LineNum, DocEntry, InvType, DocLine, InstId) " & _
							"values(@LogNum, @LineNum, @DocEntry, Convert(nvarchar(50),@DocType), 0, @InstID) End " & _
							"update R3_ObsCommon..PMT2 set SumApplied = Case @FC When 'N' Then " & getNumeric(Value) & " End, " & _
							"AppliedFC = Case @FC When 'Y' Then " & getNumeric(Value) & " End where LogNum = @LogNum and InvType = Convert(nvarchar(50),@DocType) and DocEntry = @DocEntry and InstId = @InstID "
		End Select
		conn.execute(sql)
	Case "GetDocTotal"
	Case Else
		If LineID = "" Then myTable = "TPMT" Else myTable = "PMT2"
	
		sql = "update R3_ObsCommon.." & myTable & " set " & Field & " = "
	
		If Value <> "" Then
			Select Case FieldType
				Case "S"
					sql  = sql & "N'" & saveHTMLDecode(Value, False) & "'"
				Case "N"
					sql = sql & getNumeric(Value)
				Case "D"
					sql = sql & "Convert(datetime,'" & SaveSqlDate(Value) & "',120)"
			End Select
		Else
			sql = sql & " NULL "
		End If
	
		sql = sql & " where LogNum = " & LogNum
		
		conn.execute(sql)
End Select


Select Case Field
	Case "GetDocTotal"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKGetPaymentTotalData" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LogNum") = Session("PayRetVal")
		cmd("@SlpCode") = Session("vendid")
		cmd("@MainCur") = myApp.MainCur
		cmd("@DirectRate") = GetYN(myApp.DirectRate)
		cmd("@SumDec") = myApp.SumDec
		set rs = cmd.execute()
		Response.Write FormatNumber(CDbl(rs(0)), myApp.SumDec)
	Case "DocCur"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKClearPaymentPayData"
		cmd.Parameters.Refresh()
		cmd("@LogNum") = Session("PayRetVal")
		cmd.execute()
		Response.Write "ok"
		
		'If myApp.MainCur <> Value Then
		'	set cmd = Server.CreateObject("ADODB.Command")
		'	cmd.ActiveConnection = connCommon
		'	cmd.CommandType = &H0004
		'	cmd.CommandText = "DBOLKGetCartCurrRate" & Session("ID")
		'	cmd.Parameters.Refresh()
		'	cmd("@LanID") = Session("LanID")
		'	cmd("@LogNum") = Session("PayRetVal")
		'	cmd("@Pay") = "Y"
		'	set rs = cmd.execute()
		'	Response.Write FormatNumber(CDbl(rs("Rate")), myApp.RateDec)
		'Else
		'	Response.Write FormatNumber(0, myApp.RateDec)
		'End If
	Case Else
		Response.Write "ok"
End Select

%>
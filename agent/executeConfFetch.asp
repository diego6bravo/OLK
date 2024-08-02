<!--#include file="myHTMLEncode.asp"-->
<!--#include file="authorizationClass.asp"-->
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

myType = Request.Form("Type")
ID = Request.Form("ID")
DirectRate = myApp.DirectRate

Select Case myType
	Case "S" 'Submit
		Note = Request.Form("Note")
		Status = CInt(Request.Form("Status"))
		
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKExecuteConfirmation" & Session("ID")
		cmd.Parameters.Refresh
		cmd("@ID") = ID
		cmd("@Status") = Status
		cmd("@UserSign") = Session("vendid")
		If Note <> "" Then cmd("@Note") = Note
		
		cmd.Execute()
		
		Response.Write "ok{S}" & cmd("@PoolNumber")
	Case "C"
		set rs = Server.CreateObject("ADODB.RecordSet")
		set rd = Server.CreateObject("ADODB.RecordSet")
		
		sql = 	"select T0.Status, T0.ErrCode, T0.ErrMessage, T0.ObjectCode, R3_ObsCommon.dbo.OBSSp_GetPoolNumber(T0.LogNum) PoolCount, T0.LogNum, " & _
				"T1.UserType, T1.RequestBranchID, T1.ExecAt, (select LanSign from OLKCommon..OLKLang where LanID = " & Session("LanID") & ") LanSign, " & _
				"(select PayLogNum from OLKCIC where InvLogNum = T0.LogNum) PayLogNum " & _
				"from R3_ObsCommon..TLOG T0 " & _
				"inner join OLKUAFControl T1 on T1.ID = " & ID & " and T1.LogNum = T0.LogNum "
		rs.open sql, conn, 2, 3

		If rs("Status") = "E" or rs("Status") = "S" Then
			If Not IsNull(rs("PayLogNum")) and rs("Status") = "S" Then
				sql = "select T0.Status, T0.ErrCode, T0.ErrMessage,  R3_ObsCommon.dbo.OBSSp_GetPoolNumber(T0.LogNum) PoolCount, (select DirectRate from OADM) DirectRate " & _
						"from R3_ObsCommon..TLOG T0 " & _
						"where T0.LogNum = " & rs("PayLogNum")
				set rd = conn.execute(sql)
				DirectRate = rd("DirectRate")
				Select Case rd("Status")
					Case "R"
						Dim myAut
						set myAut = New clsAuthorization
						
						sql = 	"declare @InvLogNum int set @InvLogNum = " & rs("LogNum") & " " & _
								"declare @PayLogNum int set @PayLogNum = " & rs("PayLogNum") & " " & _
								"declare @DocEntry int set @DocEntry = (select objectcode from r3_obscommon..tlog where lognum = @InvLogNum) " & _
								"declare @SumApplied numeric(19,6) set @SumApplied = (select DocTotal from oinv where docentry = @DocEntry) " & _
								"declare @Pagado numeric(19,6) set @Pagado =  " & _
								"(select IsNull(CashSum,0)+ IsNull(TrsfrSum,0) + " & _
								"(select IsNull(Sum(CheckSum),0) from r3_obscommon..pmt1 where lognum = T0.LogNum)+ " & _
								"(select IsNull(Sum(CreditSum),0) from r3_obscommon..pmt3 where lognum = T0.LogNum) " & _
								"from R3_ObsCommon..TPMT T0 " & _
								"where T0.LogNum = @PayLogNum) " & _
								"If (select DocCur from R3_ObsCommon..TPMT where LogNum = @PayLogNum) <> " & _
								"   (select DocCur from R3_ObsCommon..TDOC where LogNum = @InvLogNum) and " & _
								"   (select Currency from OCRD where CardCode = (select CardCode collate database_default from R3_ObsCommon..TDOC where LogNum = @InvLogNum)) = '##' begin " & _
								"	declare @CurRate numeric(19,6) set @CurRate = (select T1.Rate " & _
								"	from R3_ObsCommon..TDOC T0 " & _
								"	inner join ORTT T1 on T1.Currency = T0.DocCur collate database_default and DateDiff(day,RateDate,DocDate) = 0 " & _
								"	where LogNum = @InvLogNum) " & _
								"	set @Pagado = @Pagado" &  GetRateFunc(2) & "@CurRate " & _
								"end " & _
								"if @SumApplied > @Pagado begin set @SumApplied = @Pagado end " & _
								"insert r3_obscommon..pmt2(LogNum, LineNum, DocEntry, SumApplied) " & _
								"Values(@PayLogNum, 0, @DocEntry, @SumApplied) " & _
								"update R3_ObsCommon..TPMT set Series = IsNull(" & myAut.GetObjectProperty(48, "S2") & ", IsNull((select OIRRSeries from OLKBranchs where branchIndex = " & Session("branch") & "),(select Series2 from OLKDocConf where ObjectCode = 48))) where LogNum = @InvLogNum "
						conn.execute(sql)
						sql = "update r3_obscommon..tlog set status = 'C', Priority = 1 where lognum = " & rs("PayLogNum")
						conn.execute(sql)
						rs(0) = "P"
					Case Else
						rs(0) = rd(0)
						If rd(0) = "E" Then
							rs(1) = rd(1)
							rs(2) = rd(2)
						End If
						rs(3) = rd(3)
				End Select
			End If
			If rs("Status") = "S" or rs("Status") = "E" Then
				sql = "update OLKUAFControl set Status = '" & rs("Status") & "' where ID = " & ID
				Select Case rs("Status")
					Case "S"
						If Left(rs("ExecAt"), 1) <> "O" Then
							sql = sql & " EXEC OLKCommon..DBOLKObjAlert" & Session("ID") & " " & rs("LogNum") & ", " & rs("RequestBranchID") & ", '" & rs("UserType") & "', '" & rs("LanSign") & "'"
						End If
					Case "E" 
						sql = sql & " update R3_ObsCommon..TLOG set Status = 'H' where LogNum = " & rs("LogNum")
				End Select
				conn.execute(sql)
			End If
		End If

		Response.Write rs(0) & "{S}" & rs(1) & "{S}" & rs(2) & "{S}" & rs(3) & "{S}" & rs(4) 
	Case "N"
		sql = "select Note from OLKUAFControl where ID = " & ID
		set rs = Server.CreateObject("ADODB.RecordSet")
		set rs = conn.execute(sql)
		Response.Write rs(0)
	Case "P"
		sql = 	"select T0.Status, OLKCommon.dbo.DBOLKDateFormat" & Session("ID") & "(T0.ConfirmDate) ConfirmDate, T0.ConfirmDate ConfirmTime, T1.SlpName ConfirmUser, T0.Note  " & _  
				"from OLKUAFControl T0 " & _  
				"left outer join OSLP T1 on T1.SlpCode = T0.ConfirmUserSign " & _  
				"where T0.ID = " & ID
		set rs = Server.CreateObject("ADODB.RecordSet")
		set rs = conn.execute(sql)
		If Not IsNull(rs(2)) Then confirmTime = FormatDateTime(rs(2), 3)
		Response.Write rs(0) & "{S}" & rs(1) & "{S}" & confirmTime & "{S}" & rs(3) & "{S}" & rs(4)
End Select
  
Function GetRateFunc(ByVal i)
	Select Case DirectRate
		Case "Y"
			Select Case i
				Case 1
					GetRateFunc = "*"
				Case 2
					GetRateFunc = "/"
			End Select
		Case "N"
			Select Case i
				Case 1
					GetRateFunc = "/"
				Case 2
					GetRateFunc = "*"
			End Select
	End Select
End Function
        
%>
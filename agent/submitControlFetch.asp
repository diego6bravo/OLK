<!--#include file="myHTMLENcode.asp"-->
<!--#include file="authorizationClass.asp"-->
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus


set rs = Server.CreateObject("ADODB.RecordSet")
set rd = Server.CreateObject("ADODB.RecordSet")

action = CInt(Request("Action"))
LogNum = Request("LogNum")
LogNumID = Request("LogNumID")

Select Case action
	Case 0
		sql = 	"select Status, ErrCode, ErrMessage, ObjectCode, R3_ObsCommon.dbo.OBSSp_GetPoolNumber(LogNum) PoolCount, " & _  
				"Case Status When 'S' Then  " & _  
				"	Case  " & _  
				"		When Draft = 'Y' Then ObjectCode " & _
				"		When Object between 13 and 23 or Object between 203 and 204 Then Convert(nvarchar(100),OLKCommon.dbo.DBOLKGetDocNum" & Session("ID") & "(Object, Convert(int,ObjectCode))) " & _  
				"		When Object in (2,4, 33, 97, 191) Then ObjectCode " & _  
				"		When Object = 24 Then IsNull((select Convert(nvarchar(100),OLKCommon.dbo.DBOLKGetDocNum" & Session("ID") & "(13, Convert(int,ObjectCode))) from R3_ObsCommon..TLOG where LogNum = (select InvLogNum from OLKCIC where PayLogNum = T0.LogNum)), ObjectCode) " & _
				"	End  " & _  
				"End endNum, T0.Draft from R3_ObsCommon..TLOG T0 where LogNum = " & LogNum
		rs.open sql, conn, 2, 3
		
		If rs(0) = "R" and LogNumID = "PayRetVal" Then
			Dim myAut
			set myAut = New clsAuthorization
			
			sql = 	"declare @InvLogNum int set @InvLogNum = " & Session("ConfRetVal") & " " & _
					"declare @PayLogNum int set @PayLogNum = " & Session("PayRetVal") & " " & _
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
					"   (select Currency from OCRD where CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "') = '##' begin " & _
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
			sql = "update r3_obscommon..tlog set status = 'C', Priority = 1 where lognum = " & Session("PayRetVal")
			conn.execute(sql)
			rs(0) = "P"
		End If
		
		Select Case rs("Status")
			Case "E"
				sql = "update R3_ObsCommon..TLOG set Status = 'R' where LogNum = " & Request("LogNum")
				conn.execute(sql)
			Case "S"
				If userType = "C" Then
					If Not Session("noLic") Then
						set oLic = server.CreateObject("TM.LicenceConnect.LicenceConnection")
						oLic.LicenceServer = licip
						oLic.LicencePort = licport
						strResp = oLic.ConfTrans(50, 1)
					End If
				End If
				If LogNumID <> "" Then
					If Session("NotifyAdd") Then
						Session("NotifyAdd") = False
					    sql = 	"declare @LanSign nvarchar(50) set @LanSign = (select LanSign from OLKCommon..OLKLang where LanID = " & Session("LanID") & ") " & _
								"EXEC OLKCommon..DBOLKObjAlert" & Session("ID") & " " & LogNum & ", " & Session("branch") & ", '" & userType & "', @LanSign"
						conn.execute(sql)
					End If
					If Session(LogNumID) <> "" Then Session("Conf" & LogNumID) = Session(LogNumID)
					Session(LogNumID) = ""
				End If
		End Select
		
		Response.Write rs(0) & "{S}" & rs(1) & "{S}" & rs(2) & "{S}" & rs(3) & "{S}" & rs(4) & "{S}" & rs(5) & "{S}" & rs(6)

	Case 1
		sql = "update R3_ObsCommon..TLOG set Status = 'C', ErrCode = null, ErrMessage = null where LogNum = " & LogNum
		conn.execute(sql)
		Response.Write "OK"
	Case 2
		sql = "declare @LogNum int set @LogNum = " & LogNum & " " & _
			"update R3_ObsCommon..TLOGControl set Background = 'Y', tag = '" & userType & "', myLng = '" & getMyLng & "', SlpCode = " & Session("vendid") & ", ConfBranch = " & Session("branch") & " where LogNum = @LogNum " & _
			"insert R3_ObsCommon..TLOGControl(LogNum, AppID, tag, UserSign, Background, myLng, SlpCode, ConfBranch) " & _
			"select @LogNum, 'TM-OLK', 'N', " & Session("vendid") & ", 'Y', '" & getMyLng & "', " & Session("vendid") & ", " & Session("branch") & " " & _
			"where not exists(select '' from R3_ObsCommon..TLOGControl where LogNum = @LogNum) "
			conn.execute(sql)
		conn.execute(sql)
		
		Session("Conf" & LogNumID) = Session(LogNumID)
		Session(LogNumID) = ""
		Response.Write "OK"
End Select

Function GetRateFunc(ByVal i)
	sql = "select DirectRate from OADM"
	set rd = conn.execute(sql)
	Select Case rd("DirectRate")
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


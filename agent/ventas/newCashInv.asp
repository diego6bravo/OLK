<%@ Language=VBScript %>
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
<!--#include file="../authorizationClass.asp"-->
<%
response.buffer = true
Dim sap1
Dim sap2
Dim sap3
Dim sap4
Dim cmd
Dim RetVal
Dim db
Dim CardName
db = Session("olkdb")
Dim obj
Dim objs
Dim contact

Dim myAut
set myAut = New clsAuthorization

set rs = Server.CreateObject("ADODB.recordset")

sqlDec = "declare @CardCode nvarchar(15) set @CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "' " & _
		"declare @UserSeries int set @UserSeries = " & myAut.GetObjectProperty(48, "S") & " " & _
		"declare @UserSeries2 int set @UserSeries2 = " & myAut.GetObjectProperty(48, "S2") & " "

sql = sqlDec & "if (select CopyLastFCRate from olkcommon) = 'Y' begin " & _
	"	insert ORTT " & _
	"	select OLKCommon.dbo.OLKGetDateOnly(getdate()), T0.CurrCode, " & _
	"	T1.Rate, T1.DataSource, T1.UserSign " & _
	"	from OCRN T0 " & _
	"	left outer join ORTT T1 on T1.Currency = T0.CurrCode and RateDate =  " & _
	"	(select Max(RateDate) from ORTT where Currency = T0.CurrCode and DateDiff(day,OLKCommon.dbo.OLKGetDateOnly(getdate()),RateDate) < 0) " & _
	"	where CurrCode <> (select top 1 MainCurncy from oadm) and not exists( " & _
	"	select 'A' from ortt where DateDiff(day,OLKCommon.dbo.OLKGetDateOnly(getdate()),RateDate) = 0 and Currency = T0.CurrCode) " & _
	"	and exists(select 'A' from ortt where Currency = T0.CurrCode) " & _
	"end "
conn.execute(sql)

sql = sqlDec & "select case when exists( " & _
	"	select 'A' " & _
	"	from OLKCUFD T0 " & _
	"	inner join CUFD T1 on T1.TableID = T0.TableID and T1.FieldID = T0.FieldID " & _
	"	where T0.TableID = 'OITM' and T0.Active = 'Y' and not exists " & _
	"	(select 'A' from R3_ObsCommon..syscolumns where id =  " & _
	"		(select id from R3_ObsCommon..sysobjects where name = 'TITM')  " & _
	"	and name =  " & _
	"	IsNull( " & _
	"		(select SDKID collate database_default from R3_ObsCommon..TCIF where CompanyDB = db_name()),'')  " & _
	"		++ T1.AliasID) " & _
	") or exists( " & _
	"	select 'A' " & _
	"	from OLKCUFD T0 " & _
	"	inner join CUFD T1 on T1.TableID = T0.TableID and T1.FieldID = T0.FieldID " & _
	"	where T0.TableID = 'ORCT' and T0.Active = 'Y' and not exists " & _
	"	(select 'A' from R3_ObsCommon..syscolumns where id =  " & _
	"		(select id from R3_ObsCommon..sysobjects where name = 'TPMT')  " & _
	"	and name =  " & _
	"	IsNull( " & _
	"		(select SDKID collate database_default from R3_ObsCommon..TCIF where CompanyDB = db_name()),'')  " & _
	"		++ T1.AliasID) " & _
	") Then 'Y' Else 'N' End RestoreUDFErr, " & _
	"Case When Not exists(select 'A' from owhs where WhsCode = (select WhsCode from olkcommon)) Then 'Y' Else 'N' End " & _
	"WhsDefErr, " & _
	"Case When Not Exists(select 'A' from nnm1 where ObjectCode = '13' and Series = " & _
	"IsNull(@UserSeries, (select Series from OLKDocConf where ObjectCode = 48))) Then 'Y' Else 'N' End SeriesErr, " & _
	"Case When Not Exists(select 'A' from nnm1 where ObjectCode = '24' and Series =  " & _
	"IsNull(@UserSeries2, (select Series2 from OLKDocConf where ObjectCode = 48))) Then 'Y' Else 'N' End Series2Err, "
	
sql = sql & "(select Case When T0.Currency <> T1.ActCurr and T1.ActCurr <> '##' Then 'Y' Else 'N' End from OCRD T0 " & _
	"inner join OACT T1 on T1.AcctCode = T0.DebPayAcct " & _
	"where T0.CardCode = @CardCode) OCRDActCurErr, " & _
	"Case When not exists(select 'A' from R3_ObsCommon..TCIF where CompanyDB = db_name() and uid is not null) Then 'Y' Else 'N' End OBServerUserErr, " & _
	"Case When (select Active from R3_ObsCommon..TCIF where CompanyDB = db_name()) <> 'Y' Then 'Y' Else 'N' End OBServerActiveErr, "
	
sql = sql & "Case When Not Exists(select 'A' from OACT where AcctCode =  " & _
	"(select CashAcct from OLKDocConf where ObjectCode = 48)) or " & _
	"Not Exists(select 'A' from OACT where AcctCode =  " & _
	"(select CheckAcct from OLKDocConf where ObjectCode = 48)) Then 'Y' Else 'N' End PayAcctErr, " & _
	"(select  " & _
	"case when T0.Currency <> (select top 1 MainCurncy from oadm)  " & _
	"and ( " & _
	"	(T0.Currency <> '##'  " & _
	"	and not exists(select 'A' from ORTT where Currency = T0.Currency and DateDiff(day,getdate(),RateDate) = 0))  " & _
	"	or  " & _
	"	(T0.Currency = '##'  " & _
	"	and exists(select T0.CurrCode from OCRN T0 " & _
	"	left outer join ORTT T1 on T1.Currency = T0.CurrCode and DateDiff(day,getdate(),RateDate) = 0 " & _
	"	where T0.CurrCode <> (select top 1 MainCurncy from oadm) and T1.Currency is null)) " & _
	") Then 'Y' Else 'N' End CurRateErr " & _
	"from ocrd T0 where CardCode = @CardCode) CurRateErr, " & _
	"Case When '" & myAut.HasAuthorization(60) & "' = 'False' Then Case When (select SlpCode from OCRD where CardCode = @CardCode) <> " & Session("VendID") & " Then 'Y' Else 'N' End End AsignedSLP "
	set rs = conn.execute(sql)
	If rs("AsignedSLP") = "Y" Then Response.Redirect "../configErr.asp&errCmd=AsignedSLP"
	For each itm in rs.Fields
		if itm = "Y" Then Response.Redirect "../configErr.asp?errCmd=PayDoc"
	next

sql = "declare @branchIndex int set @branchIndex = " & Session("branch") & " " & _
"select CardCode, IsNull(Replace(cardname,'''', ''''''),'') CardName,  " & _
"Case Currency When '##' Then (select top 1 MainCurncy from oadm) Else Currency End Currency, T0.Address,  " & _
"(select cntctcode from ocpr where cardcode = T0.cardcode and name = T0.cntctprsn) CntctCode, OLKCommon.dbo.DBOLKGetCardPList" & Session("ID") & "(N'" & saveHTMLDecode(Session("UserName"), False) & "', '" & userType & "') listnum, T0.groupnum, " & _
"IsNull(" & myAut.GetObjectProperty(48, "S") & ", IsNull((select OIRISeries from OLKBranchs where branchIndex = @branchIndex),T2.Series)) OINVSeries, " & _
"IsNull(" & myAut.GetObjectProperty(48, "S2") & ", IsNull((select OIRRSeries from OLKBranchs where branchIndex = @branchIndex),T2.Series2)) ORCTSeries, " & _
"IsNull((select OIRCashAcct from OLKBranchs where branchIndex = @branchIndex),T2.CashAcct) CashAcct, " & _
"IsNull((select OIRCheckAcct from OLKBranchs where branchIndex = @branchIndex),T2.CheckAcct) CheckAcct, " & _
"(select cntctcode from ocpr where cardcode = T0.cardcode and name = T0.cntctprsn) CntctCode, " & _
"DateAdd(day,ExtraDays,DateAdd(month,ExtraMonth, " & _
"Case PayDuMonth " & _
"	When 'N' Then getdate()  " & _
"	When 'Y' Then DateAdd(day,1-day(DateAdd(month,1,getdate())),DateAdd(month,1,getdate())) " & _
"	When 'H' Then DateAdd(day,15-day(DateAdd(month,1,getdate())),DateAdd(month,1,getdate())) " & _
"	When 'E' Then DateAdd(day,1-day(DateAdd(month,1,getdate())),DateAdd(month,1,getdate()))-1 " & _
"End)) DocDueDate " & _
"from ocrd T0 " & _
"cross join olkcommon T1 " & _
"cross join OLKDocConf T2 " & _
"inner join octg T3 on T3.GroupNum = T0.GroupNum " & _
"where cardcode = N'" & saveHTMLDecode(Session("UserName"), False) & "' and T2.ObjectCode = 48 "

set rs = conn.execute(sql)
Session("PriceList") = RS("listnum")
If IsNull(rs("CntctCode")) Then cntctcode = "NULL" else cntctcode = rs("cntctcode")
           
set conn2=Server.CreateObject("ADODB.Connection")
conn2.Provider=olkSqlProv
conn2.Open  "Provider=SQLOLEDB;charset=utf8;" & _
          "Data Source=" & olkip & ";" & _
          "Initial Catalog=R3_ObsCommon;" & _
          "Uid=" & olklogin & ";" & _
          "Pwd=" & olkpass & ""
set Cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = conn2
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "OBSSp_Request"
cmd.Parameters.Refresh
cmd.Execute , Array(0, db, Null, 13, "A", Null)
RetVal = cmd.Parameters.Item(0).Value
Session("RetVal") = RetVal

cmd.Execute , Array(0, db, Null, 24, "A", Null)
PayRetVal = cmd.Parameters.Item(0).Value
Session("PayRetVal") = PayRetVal
If rs("CntctCode") <> "" Then CntctCode = rs("CntctCode") Else CntctCode = "NULL"
sql = 	"INSERT INTO TDOC(LogNum, CardCode, CardName, DocDate, DocDueDate, CntctCode, Series, SLPCode, GroupNum, DocCur) " & _
		"VALUES(" & RetVal & ", N'" & saveHTMLDecode(Session("UserName"), False) & "', N'" & rs("cardname") & "', getdate(), " & _
		"'" & SaveSqlDate(FormatDate(rs("DocDueDate"), False)) & "', " & CntctCode & ", '" & rs("OINVSeries") & "', '" & Session("vendid") & "', " & rs("GroupNum") & ", N'" & rs("Currency") & "') " & _
		"insert TLOGControl(LogNum, appId, tag, UserSign, PriceList) values(" & RetVal & ", 'TM-OLK', 'N', " & Session("vendid") & ", " & Session("PriceList") & ") "
conn2.execute(sql)

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKPostObjectCreation" & Session("ID")
cmd.Parameters.Refresh
Select Case userType
	Case "C"
		cmd("@sessiontype") = "C"
	Case "V"
		cmd("@sessiontype") = "A"
End Select
cmd("@transtype") = "N"
cmd("@object") = 48 
cmd("@LogNum") = RetVal
cmd("@CurrentSlpCode") = Session("vendid")
cmd("@Branch") = Session("branch")
cmd.execute()

If myApp.LawsSet = "MX" or myApp.LawsSet = "CL" or myApp.LawsSet = "CR" or myApp.LawsSet = "GT" or myApp.LawsSet = "US" or myApp.LawsSet = "CA" Then
	sql = "update R3_ObsCommon..tdoc set ShipToCode = (select ShipToDef from ocrd where CardCode = T0.CardCode collate database_default) " & _
	"from R3_ObsCommon..tdoc T0 where LogNum = " & RetVal
	conn.execute(sql)
End If
			sql = "insert tpmt(LogNum, DocType, DocDate, CardCode, CardName, Address, DocCur, CntctCode, Series, JrnlMemo, CashAcct, CheckAcct) " & _
				  "Values(" & PayRetVal & ",'C',getdate(), N'" & rs("CardCode") & "', N'" & rs("CardName") & "', N'" & rs("Address") & _
				  "', N'" & rs("Currency") & "', " & CntctCode & ", " & rs("ORCTSeries") & ", N'Recibo - " & rs("CardCode") & "', N'" & rs("CashAcct") & "', N'" & rs("CheckAcct") & "')"
			'response.redirect "http://www.topmanage.com.pa/query.asp?query=" & sql
			conn2.execute(sql)
			sql = "insert OLKDocControl(slpcode, lognum, DocType) Values(" & Session("vendid") & ", " & PayRetval & ", 13) " & _
			"insert OLKCic(InvLogNum, PayLogNum) values(" & Session("RetVal") & ", " & Session("PayRetVal") & ")"
			conn.execute(sql)
set rs = nothing
conn2.close
conn.close
Session("cart") = "cart"
Session("PayCart") = True
Response.Redirect "../cart.asp"

%>
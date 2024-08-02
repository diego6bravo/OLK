<% addLngPathStr = "" %>
<!--#include file="lang/cxcData.asp" -->
<head>
<style type="text/css">
.style1 {
	border-width: 0px;
}
.style2 {
	text-align: center;
}
</style>
</head>
<% If Request("excell") <> "Y" Then %>
<script type="text/javascript" src="scr/calendar.js"></script>
<script type="text/javascript" src="scr/lang/calendar-<%=Left(Session("myLng"), 2)%>.js"></script>
<script type="text/javascript" src="scr/calendar-setup.js"></script>
<% End If %>
<% If (Request("LinkRep") = "Y" and (Request("Excell") <> "Y" and Request("PDF") <> "Y")) or userType = "C" Then %>
<div id="tblSave" style="text-align: right">
<a href="#" onclick="javascript:printStory();"><img alt="<%=getcxcDataLngStr("DtxtPrint")%>" border="0" src="images/print_OLK.gif"></a>&nbsp;
<a href="#" onclick="javascript:doCxcExport('pdf');"><img alt="<%=getcxcDataLngStr("DtxtExpPDF")%>" border="0" src="images/pdf_OLK.gif"></a>&nbsp;
<a href="#" onclick="javascript:doCxcExport('excell');"><img alt="<%=getcxcDataLngStr("DtxtExpToExcell")%>" border="0" src="images/excell.gif"></a>
</div>
<script type="text/javascript" src="cxc.js"></script><% End If
set ra = Server.CreateObject("ADODB.RecordSet")
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetObjectData" & Session("ID")
cmd.Parameters.Refresh()
cmd("@ObjType") = "S"
cmd("@ObjID") = 12
cmd("@UserType") = userType
set ra = cmd.execute()
If ra("CustomTitle") = "Y" Then
	strTitle = ra("ObjContent")
	strTitle = Replace(strTitle, "{SelDes}", SelDes)
	strTitle = Replace(strTitle, "{rtl}", Session("rtl"))  
	strTitle = Replace(strTitle, "{CmpName}", CmpName)
	%><input type="hidden" id="hdTitle" value="<%=Server.HTMLEncode(strTitle)%>"><%
End If
If Session("UserName") = "" and userType = "V" Then response.redirect "searchClient.asp"
If Request("excell") = "Y" and Request("pdf") <> "Y" then 
	response.ContentType="application/vnd.ms-excel" %>
<!--#include file="myHTMLEncode.asp"-->
<% End If 
If Request("LinkRep") <> "Y" Then
	If userType = "C" Then MainDoc = "default.asp" Else MainDoc = "agent.asp"
Else
	MainDoc = "cxcPrint.asp"
End If

varx = 0
           set rs = Server.CreateObject("ADODB.recordset")
           set rd = Server.CreateObject("ADODB.recordset")
           set rc = Server.CreateObject("ADODB.recordset")
           set rxy1 = Server.CreateObject("ADODB.recordset")
sqlAdd = ""
If Request("CurrCode") = "" or Request("CurrCode") = "F" Then
	sqlAdd = "Case When Currency = '##' Then (select top 1 MainCurncy from oadm) Else Currency End Currency, "
ElseIf Request("CurrCode") = "L" Then
	sqlAdd = "(select top 1 MainCurncy from oadm) Currency, "
End If
sql = "select IsNull(CardName, '') CardName, " & sqlAdd & " getdate() As CDate, " & _
		"convert(datetime, convert(int,getdate())-ecdays) as 'Date', (select top 1 DispPosDeb from OADM) DispPosDeb, (select top 1 MainCurncy from oadm) MainCur " & _
	  "from olkcommon " & _
	  "cross join ocrd where cardcode = N'" & saveHTMLDecode(Session("UserName"), False) & "'"
set rs = conn.execute(sql)
CardName = rs("CardName")
MainCur = rs("MainCur")
BPCur = rs("Currency")
If Request("timestamp1") <> "" Then vardate1 = Request("timestamp1") else vardate1 = FormatDate(rs("Date"), False)
If Request("timestamp2") <> "" Then vardate2 = Request("timestamp2") else vardate2 = FormatDate(rs("CDate"), False)

If rs("DispPosDeb") = "Y" Then
	DispPosDeb = 1
Else
	DispPosDeb = -1
End If

If myApp.GetShowCxcOpenInvBy = "DocDate" Then jdtDateFilter = "RefDate" Else jdtDateFilter = "DueDate"

sql = "declare @MainCur nvarchar(3) set @MainCur = (select top 1 MainCurncy from oadm) " & _
	  "select (select isnull(sum(debit - credit),0) from jdt1 T0 " & _
	  "where shortname = N'" & saveHTMLDecode(Session("UserName"), False) & "' And " & jdtDateFilter & " < Convert(datetime,'" & SaveSqlDate(vardate1) & "',120)) As oldbalance, " & _
	  "Case When @MainCur = N'" & rs("Currency") & "' Then 1 Else (" & _
	  "select top 1 (Rate) from ortt where DateDiff(day,RateDate,getdate()) >= 0 and Currency = N'" & rs("Currency") & "' order by RateDate desc " & _
	  ") End CurRate "
set rd = conn.execute(sql)
varx = CDbl(rd("oldbalance"))
CurRate = CDbl(rd("CurRate"))

sql = "select RefDate, DueDate, " & _
	  "transtype, Case When ref1 is not null and ref1 <> '' Then ref1 Else Convert(nvarchar(20),TransID) End ref1, CreatedBy, IsNull(linememo, '') linememo, debit, credit ,debit - credit as tempsaldo " & _
	  "from jdt1 where ShortName = N'" & saveHTMLDecode(Session("UserName"), False) & "' and DateDiff(day, " & jdtDateFilter & ", Convert(datetime,'" & SaveSqlDate(vardate1) & "',120)) <= 0 " & _
	  "and DateDiff(day," & jdtDateFilter & ", Convert(datetime,'" & SaveSqlDate(vardate2) & "',120)) >= 0 " & _
	  "order by "
	  
If myApp.GetShowCxcOpenInvBy = "DocDueDate" Then sql = sql & "DueDate, "

sql = sql & "refdate, TransId"
set rc = conn.execute(sql)

If myApp.SVer >= 8 Then
	colCredit = "BalDueCred"
	colDebit = "BalDueDeb"
Else
	colCredit = "Credit"
	colDebit = "Debit"
End If

sql = "declare @varDate datetime set @varDate = Convert(datetime,'" & SaveSqlDate(vardate2) & "',120) " & _
"declare @CardCode nvarchar(15) declare @Credit numeric(19,6) declare @d121 numeric(19,6) declare @d120 numeric(19,6) declare @d90 numeric(19,6) declare @d60 numeric(19,6) declare @d30 numeric(19,6) " & _
"set @CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "'  " & _
"set @Credit = (select isnull(sum(" & colCredit & "),0) From jdt1 where shortname = @CardCode and " & jdtDateFilter & " <= @varDate) +(select isnull(sum(" & colDebit & "*-1),0) From jdt1 where shortname = @CardCode and " & jdtDateFilter & " <= getdate() and TransType = 14)  " & _
"set @d121 = (select isnull(sum(" & colDebit & "),0) From jdt1 where shortname = @CardCode and DateDiff(day," & jdtDateFilter & ",@varDate) >= 121 and TransType <> 14)  " & _
"set @d120 = (select isnull(sum(" & colDebit & "),0) From jdt1 where shortname = @CardCode and DateDiff(day," & jdtDateFilter & ",@varDate) Between 91 and 120 and TransType <> 14)  " & _
"set @d90 = (select isnull(sum(" & colDebit & "),0) From jdt1 where shortname = @CardCode and DateDiff(day," & jdtDateFilter & ",@varDate) Between 61 and 90 and TransType <> 14)  " & _
"set @d60 = (select isnull(sum(" & colDebit & "),0) From jdt1 where shortname = @CardCode and DateDiff(day," & jdtDateFilter & ",@varDate) Between 31 and 60 and TransType <> 14)  " & _
"set @d30 = (select isnull(sum(" & colDebit & "),0) From jdt1 where shortname = @CardCode and DateDiff(day," & jdtDateFilter & ",@varDate) between 0 and 30 and TransType <> 14)  " & _
"set @d121 = @d121 - @credit  " & _
"	If @d121 < 0 Begin  " & _
"		set @credit = @d121  " & _
"		set @d121 = 0 End  " & _
"	Else Begin  " & _
"		set @Credit = 0 End  " & _
"set @d120 = @d120 + @credit  " & _
"	If @d120 < 0 Begin  " & _
"		set @Credit = @d120  " & _
"		set @d120 = 0 End  " & _
"	Else Begin  " & _
"		set @Credit = 0 End  " & _
"set @d90 = @d90 + @credit  " & _
"	If @d90 < 0 Begin  " & _
"		set @Credit = @d90  " & _
"		set @d90 = 0 End  " & _
"	Else Begin  " & _
"		set @Credit = 0 End  " & _
"set @d60 = @d60 + @credit  " & _
"	If @d60 < 0 Begin  " & _
"		set @Credit = @d60  " & _
"		set @d60 = 0 End  " & _
"	Else Begin  " & _
"		set @Credit = 0 End  " & _
"set @d30 = @d30 + @credit  " & _
"select @d121'121+', @d120 '120', @d90 '90', @d60 '60', @d30 '30'"

If jdtDateFilter = "DueDate" Then
	sql = sql & ", (select IsNull(Sum(" & colDebit & "),0)-IsNull(Sum(" & colCredit & "),0) from jdt1 where ShortName = @CardCode and DueDate > @varDate) Futuro "
End If

sql = sql & " from ocrd where CardCode = @CardCode  "
set rxy1 = conn.execute(sql)
d121 = CDBl(rxy1("121+"))
d120 = CDBl(rxy1("120"))
d90 = CDBl(rxy1("90"))
d60 = CDBl(rxy1("60"))
d30 = CDBl(rxy1("30"))

sql = "declare @ObjType char(1) set @ObjType = 'S' " & _
"declare @ObjId int set @ObjId = 12 " & _
"if exists(select 'A' from OLKObjects where ObjType = @ObjType and ObjId = @ObjId and Status = 'Y' and '" & userType & "' = 'C') begin " & _
"	select ObjContent " & _
"	from OLKObjects where ObjType = @ObjType and ObjId = @ObjId " & _
"end else begin " & _
"	select ObjContent " & _
"	from OLKCommon..OLKObjects where ObjType = @ObjType and ObjId = @ObjId " & _
"end "
set ra = conn.execute(sql)
strContent = ra("ObjContent")
strContent = Replace(strContent, "{SelDes}", SelDes)
strContent = Replace(strContent, "{rtl}", Session("rtl"))
strContent = Replace(strContent, "{cmpName}", cmpName)
strContent = Replace(strContent, "{vardate1}", vardate1)
strContent = Replace(strContent, "{vardate2}", vardate2)
strContent = Replace(strContent, "{LtxtMy}", getcxcDataLngStr("LtxtMy"))
strContent = Replace(strContent, "{txtCXC}", txtCXC)
strContent = Replace(strContent, "{LtxtLocalCur}", getcxcDataLngStr("LtxtLocalCur"))
strContent = Replace(strContent, "{LocalCur}", MainCur)
strContent = Replace(strContent, "{LtxtCardCurrency}", getcxcDataLngStr("LtxtCardCurrency"))
strContent = Replace(strContent, "{BPCur}", BPCur)
If MainCur = BPCur Then
	strContent = Replace(strContent, "{HideCurSel}", "style=""display: none; """)
Else
		strContent = Replace(strContent, "{HideCurSel}", "")
End If
'strContent = Replace(strContent, "{LbtnPendDocs}", getcxcDataLngStr("LbtnPendDocs"))
strContent = Replace(strContent, "{DtxtQuery}", getcxcDataLngStr("DtxtQuery"))
strContent = Replace(strContent, "{DtxtDate}", getcxcDataLngStr("DtxtDate"))
strContent = Replace(strContent, "{DtxtDueDate}", getcxcDataLngStr("DtxtDueDate"))
strContent = Replace(strContent, "{DtxtType}", getcxcDataLngStr("DtxtType"))
strContent = Replace(strContent, "{LtxtRef}", getcxcDataLngStr("LtxtRef"))
strContent = Replace(strContent, "{LtxtMemo}", getcxcDataLngStr("LtxtMemo"))
strContent = Replace(strContent, "{LtxtDebit}", getcxcDataLngStr("LtxtDebit"))
strContent = Replace(strContent, "{LtxtCredit}", getcxcDataLngStr("LtxtCredit"))
strContent = Replace(strContent, "{DtxtBalance}", getcxcDataLngStr("DtxtBalance"))
strContent = Replace(strContent, "{LtxtDocNum}", getcxcDataLngStr("LtxtDocNum"))
strContent = Replace(strContent, "{LtxtCondition}", getcxcDataLngStr("LtxtCondition"))
strContent = Replace(strContent, "{DtxtImport}", getcxcDataLngStr("DtxtImport2"))
strContent = Replace(strContent, "{LtxtDays}", getcxcDataLngStr("LtxtDays"))
strContent = Replace(strContent, "{LtxtPerInitBal}", getcxcDataLngStr("LtxtPerInitBal"))
strContent = Replace(strContent, "{LtxtPerEndBal}", getcxcDataLngStr("LtxtPerEndBal"))
strContent = Replace(strContent, "{LtxtBalToDate}", getcxcDataLngStr("LtxtBalToDate"))
strContent = Replace(strContent, "{LtxtFuture}", getcxcDataLngStr("LtxtFuture"))
strContent = Replace(strContent, "{LtxtDateRange}", getcxcDataLngStr("LtxtDateRange"))
strContent = Replace(strContent, "{LttlPendInv}", Replace(getcxcDataLngStr("LttlPendInv"), "{0}", Server.HTMLEncode(txtInvs)))
strContent = Replace(strContent, "{MainDoc}", MainDoc)
If InStr(strContent, "startDueDate1") <> 0 Then
	If GetShowCxcDueDate Then
		strContent = Replace(strContent, "{PerEndBalColSpan}", 6)
		For i = 1 to 6
			strContent = Replace(strContent, getFullMid(strContent, "startDueDate" & i, "endDueDate" & i), getMid(strContent, "startDueDate" & i, "endDueDate" & i))
		Next
	Else
		strContent = Replace(strContent, "{PerEndBalColSpan}", 5)
		For i = 1 to 6
			strContent = Replace(strContent, getFullMid(strContent, "startDueDate" & i, "endDueDate" & i), "")
		Next
	End If
End If
If Session("rtl") = "" Then
	strContent = Replace(strContent, "{rtl}", "")
Else
	strContent = Replace(strContent, "{rtl}", "rtl/")
End If

If Request("excell") <> "Y" Then
	strContent = Replace(strContent, getFullMid(strContent, "startExcel", "endExcel"), "")
	strContent = Replace(strContent, getFullMid(strContent, "startNoExcel", "endNoExcel"), getMid(strContent, "startNoExcel", "endNoExcel"))
	If Request("CurrCode") = "F" or Request("CurrCode") = "" Then
		strContent = Replace(strContent, "{selectedcurf}", "selected")
	Else
		strContent = Replace(strContent, "{selectedcurf}", "")
	End If
	
	If InStr(strContent, "startNoExcel2") <> 0 Then
		strContent = Replace(strContent, getFullMid(strContent, "startNoExcel2", "endNoExcel2"), getMid(strContent, "startNoExcel2", "endNoExcel2"))
	End If
Else
	strContent = Replace(strContent, getFullMid(strContent, "startNoExcel", "endNoExcel"), "")
	strContent = Replace(strContent, getFullMid(strContent, "startExcel", "endExcel"), getMid(strContent, "startExcel", "endExcel"))
	If InStr(strContent, "startNoExcel2") <> 0 Then
		strContent = Replace(strContent, getFullMid(strContent, "startNoExcel2", "endNoExcel2"), "")
	End If
End If

InitBal = CDbl(rd("oldbalance"))
If DirectRate = "Y" Then InitBal = InitBal/CurRate Else InitBal = InitBal*CurRate
strContent = Replace(strContent, "{InitCurr}", rs("Currency"))
strContent = Replace(strContent, "{InitBal}", FormatNumber(InitBal*DispPosDeb,myApp.SumDec))
strContent = Replace(strContent, "{CardName}", CardName)
strContent = Replace(strContent, "{CardCode}", Session("UserName"))
If userType = "C" Then
	If InStr(strContent, "startCrdLnk") <> 0 Then
		strContent = Replace(strContent, getFullMid(strContent, "startCrdLnk", "endCrdLnk"), "")
	End If
Else
	If Request("excell") <> "Y" Then
		strContent = Replace(strContent, getFullMid(strContent, "startCrdLnk", "endCrdLnk"), getMid(strContent, "startCrdLnk", "endCrdLnk"))
	Else
		strContent = Replace(strContent, getFullMid(strContent, "startCrdLnk", "endCrdLnk"), "")
	End If
End If

strLoopEnd = ""
do while not rc.eof
	tmpStr = getMid(strContent, "startLoopJdt", "endLoopJdt")
	
	tmpStr = Replace(tmpStr, "{Date}", FormatDate(RC("RefDate"), True))
	tmpStr = Replace(tmpStr, "{DueDate}", FormatDate(RC("DueDate"), True))

	If Request("excell") <> "Y" and (Rc("transtype") = "13" or Rc("transtype") = "14" or rc("TransType") = "24" or rc("TransType") = "46" or rc("TransType") = "18" or rc("TransType") = "19") Then
		tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startJdtDetLnk", "endJdtDetLnk"), getMid(tmpStr, "startJdtDetLnk", "endJdtDetLnk"))
		tmpStr = Replace(tmpStr, "{TransType}", rc("TransType"))
		tmpStr = Replace(tmpStr, "{CreatedBy}", rc("CreatedBy"))
	Else
		tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startJdtDetLnk", "endJdtDetLnk"), "")
	End If
	Select Case CStr(Rc("transtype"))
		Case "-2"
			tmpStr = Replace(tmpStr, "{transtype}", getcxcDataLngStr("LtxtTransTypeOB"))
		Case "13"
			tmpStr = Replace(tmpStr, "{transtype}", getcxcDataLngStr("LtxtTransTypeIN"))
		Case "24" 
			tmpStr = Replace(tmpStr, "{transtype}", getcxcDataLngStr("LtxtTransTypeRC"))
		Case "14" 
			tmpStr = Replace(tmpStr, "{transtype}", getcxcDataLngStr("LtxtTransTypeCN"))
		Case "30" 
			tmpStr = Replace(tmpStr, "{transtype}", getcxcDataLngStr("LtxtTransTypeJE"))
		Case "57" 
			tmpStr = Replace(tmpStr, "{transtype}", getcxcDataLngStr("LtxtTransTypeCP"))
		Case "46" 
			tmpStr = Replace(tmpStr, "{transtype}", getcxcDataLngStr("LtxtTransTypePS"))
		Case "18" 
			tmpStr = Replace(tmpStr, "{transtype}", getcxcDataLngStr("LtxtTransTypePI"))
		Case "19" 
			tmpStr = Replace(tmpStr, "{transtype}", getcxcDataLngStr("LtxtTransTypeDN"))
		Case Else
			tmpStr = Replace(tmpStr, "{transtype}", "")
	End Select
	tmpStr = Replace(tmpStr, "{ref1}", RC("ref1"))
	tmpStr = Replace(tmpStr, "{linememo}", RC("linememo"))
	
	If Rc("debit") <> "0" or rc("TransType") = "13" Then
		curDebit = CDbl(rc("debit"))
		If DirectRate = "Y" Then curDebit = curDebit/CurRate Else curDebit = curDebit*CurRate
		If Rc("debit") < "0" Then
			tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startJdtDebitNeg", "endJdtDebitNeg"), getMid(tmpStr, "startJdtDebitNeg", "endJdtDebitNeg"))
			tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startJdtDebit", "endJdtDebit"), "")
			tmpStr = Replace(tmpStr, "{curDebit}", rs("Currency") & "&nbsp;" & FormatNumber(curDebit*-1,myApp.SumDec))
		Else
			tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startJdtDebitNeg", "endJdtDebitNeg"), "")
			tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startJdtDebit", "endJdtDebit"), getMid(tmpStr, "startJdtDebit", "endJdtDebit"))
			tmpStr = Replace(tmpStr, "{curDebit}", rs("Currency") & "&nbsp;" & FormatNumber(curDebit,myApp.SumDec))
		End If
	Else
		tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startJdtDebitNeg", "endJdtDebit"), "")
	End If
	
	If Rc("credit") <> "0" Then
		curCredit = CDbl(rc("credit"))
		If DirectRate = "Y" Then curCredit = curCredit/CurRate Else curCredit = curCredit*CurRate
		If Rc("credit") < "0" Then
			tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startJdtCreditNeg", "endJdtCreditNeg"), getMid(tmpStr, "startJdtCreditNeg", "endJdtCreditNeg"))
			tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startJdtCredit", "endJdtCredit"), "")
			tmpStr = Replace(tmpStr, "{curCredit}", rs("Currency") & "&nbsp;" & FormatNumber(curCredit*-1,myApp.SumDec))
		Else
			tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startJdtCreditNeg", "endJdtCreditNeg"), "")
			tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startJdtCredit", "endJdtCredit"), getMid(tmpStr, "startJdtCredit", "endJdtCredit"))
			tmpStr = Replace(tmpStr, "{curCredit}", rs("Currency") & "&nbsp;" & FormatNumber(curCredit,myApp.SumDec))
		End If
	Else
		tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startJdtCreditNeg", "endJdtCredit"), "")
	End If
	
	varx = CDbl(rc("tempsaldo")) + varx
	If DirectRate = "Y" Then tempSaldo = CDbl(varx)/CurRate Else tempSaldo = CDbl(varx)*CurRate
	tempSaldo=tempSaldo*DispPosDeb
	If tempSaldo < 0 then
			tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startJdtBalNeg", "endJdtBalNeg"), getMid(tmpStr, "startJdtBalNeg", "endJdtBalNeg"))
			tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startJdtBal", "endJdtBal"), "")
			tmpStr = Replace(tmpStr, "{tempSaldo}", rs("Currency") & "&nbsp;" & FormatNumber(tempSaldo*-1,myApp.SumDec))
	Else
			tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startJdtBalNeg", "endJdtBalNeg"), "")
			tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startJdtBal", "endJdtBal"), getMid(tmpStr, "startJdtBal", "endJdtBal"))
			tmpStr = Replace(tmpStr, "{tempSaldo}", rs("Currency") & "&nbsp;" & FormatNumber(tempSaldo,myApp.SumDec))
	End If
	If DispPosDeb = -1 Then tempSaldo = tempSaldo*-1
	
	strLoopEnd = strLoopEnd & tmpStr
rc.movenext
loop

strContent = Replace(strContent, getFullMid(strContent, "startLoopJdt", "endLoopJdt"), strLoopEnd)

If DirectRate = "Y" Then varx = varx/CurRate Else varx = varx*CurRate
varx = varx * DispPosDeb
If varx < 0 then
		strContent = Replace(strContent, getFullMid(strContent, "startEndBalNeg", "endEndBalNeg"), getMid(strContent, "startEndBalNeg", "endEndBalNeg"))
		strContent = Replace(strContent, getFullMid(strContent, "startEndBal", "endEndBal"), "")
		strContent = Replace(strContent, "{varx}", rs("Currency") & "&nbsp;" & FormatNumber(varx*-1,myApp.SumDec))
Else
		strContent = Replace(strContent, getFullMid(strContent, "startEndBalNeg", "endEndBalNeg"), "")
		strContent = Replace(strContent, getFullMid(strContent, "startEndBal", "endEndBal"), getMid(strContent, "startEndBal", "endEndBal"))
		strContent = Replace(strContent, "{varx}", rs("Currency") & "&nbsp;" & FormatNumber(varx,myApp.SumDec))
End If
If DispPosDeb = -1 Then varx = varx * -1

If jdtDateFilter = "DueDate" Then
	strContent = Replace(strContent, getFullMid(strContent, "startShowFuture1", "endShowFuture1"), getMid(strContent, "startShowFuture1", "endShowFuture1"))
	strContent = Replace(strContent, getFullMid(strContent, "startShowFuture2", "endShowFuture2"), getMid(strContent, "startShowFuture2", "endShowFuture2"))
	strContent = Replace(strContent, "{Futuro}", rs("Currency") & "&nbsp;" & FormatNumber(CDbl(rxy1("Futuro"))*DispPosDeb,myApp.SumDec))
Else
	strContent = Replace(strContent, getFullMid(strContent, "startShowFuture1", "endShowFuture1"), "")
	strContent = Replace(strContent, getFullMid(strContent, "startShowFuture2", "endShowFuture2"), "")
End If

If DirectRate = "Y" Then
	d30 = d30/CurRate
	d60 = d60/CurRate
	d90 = d90/CurRate
	d120 = d120/CurRate
	d121 = d121/CurRate
ElseIf DirectRate = "N" Then
	d30 = d30*CurRate
	d60 = d60*CurRate
	d90 = d90*CurRate
	d120 = d120*CurRate
	d121 = d121*CurRate
End If

d30 = d30 * DispPosDeb
d60 = d60 * DispPosDeb
d90 = d90 * DispPosDeb
d120 = d120 * DispPosDeb
d121 = d121 * DispPosDeb

If d30 < 0 Then
	strContent = Replace(strContent, getFullMid(strContent, "startD30Neg", "endD30Neg"), getMid(strContent, "startD30Neg", "endD30Neg"))
	strContent = Replace(strContent, getFullMid(strContent, "startD30", "endD30"), "")
	strContent = Replace(strContent, "{d30}", rs("Currency") & "&nbsp;" & FormatNumber(d30*-1,myApp.SumDec))
Else
	strContent = Replace(strContent, getFullMid(strContent, "startD30Neg", "endD30Neg"), "")
	strContent = Replace(strContent, getFullMid(strContent, "startD30", "endD30"), getMid(strContent, "startD30", "endD30"))
	strContent = Replace(strContent, "{d30}", rs("Currency") & "&nbsp;" & FormatNumber(d30,myApp.SumDec))
End If

strContent = Replace(strContent, "{d60}", rs("Currency") & "&nbsp;" & FormatNumber(d60,myApp.SumDec))
strContent = Replace(strContent, "{d90}", rs("Currency") & "&nbsp;" & FormatNumber(d90,myApp.SumDec))
strContent = Replace(strContent, "{d120}", rs("Currency") & "&nbsp;" & FormatNumber(d120,myApp.SumDec))
strContent = Replace(strContent, "{d121}", rs("Currency") & "&nbsp;" & FormatNumber(d121,myApp.SumDec))

If myApp.GetShowCxcOpenInvBy = "DocDueDate" and myApp.GetShowCxcIncTrans Then
	sql = "select RefDate, DueDate, " & _
	  "transtype, Case When ref1 is not null and ref1 <> '' Then ref1 Else Convert(nvarchar(20),TransID) End ref1, CreatedBy, IsNull(linememo, '') linememo, debit, credit ,debit - credit as tempsaldo, Convert(int,DueDate) DueDateInt " & _
	  "from jdt1 where ShortName = N'" & saveHTMLDecode(Session("UserName"), False) & "' and DateDiff(day," & jdtDateFilter & ",Convert(datetime,'" & SaveSqlDate(vardate2) & "',120)) < 0 " & _
	  "order by DueDateInt asc, TransId"
	set rc = conn.execute(sql)
	If Not rc.Eof Then
		strContent = Replace(strContent, getFullMid(strContent, "startIncTrans", "endIncTrans"), getMid(strContent, "startIncTrans", "endIncTrans"))
		strContent = Replace(strContent, "{LtxtIncTrans}", getcxcDataLngStr("LtxtIncTrans"))
		strLoopEnd = ""
		do while not rc.eof
			tmpStr = getMid(strContent, "startIncTransLoop", "endIncTransLoop")
			tmpStr = Replace(tmpStr, "{Date}", FormatDate(rc("RefDate"), True))
			tmpStr = Replace(tmpStr, "{DueDate}", FormatDate(rc("DueDate"), True))
			tmpStr = Replace(tmpStr, "{TransType}", rc("transtype"))
			
			If Request("excell") <> "Y" and (Rc("transtype") = "13" or Rc("transtype") = "14" or rc("TransType") = "24" or rc("TransType") = "46" or rc("TransType") = "18" or rc("TransType") = "19") Then
				tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startIncTransLink", "endIncTransLink"), getMid(tmpStr, "startIncTransLink", "endIncTransLink"))
				tmpStr = Replace(tmpStr, "{TransType}", rc("TransType"))
				tmpStr = Replace(tmpStr, "{CreatedBy}", rc("CreatedBy"))
			Else
				tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startIncTransLink", "endIncTransLink"), "")
			End If

			Select Case CStr(Rc("transtype"))
				Case "-2"
					tmpStr = Replace(tmpStr, "{transtype}", getcxcDataLngStr("LtxtTransTypeOB"))
				Case "13"
					tmpStr = Replace(tmpStr, "{transtype}", getcxcDataLngStr("LtxtTransTypeIN"))
				Case "24" 
					tmpStr = Replace(tmpStr, "{transtype}", getcxcDataLngStr("LtxtTransTypeRC"))
				Case "14" 
					tmpStr = Replace(tmpStr, "{transtype}", getcxcDataLngStr("LtxtTransTypeCN"))
				Case "30" 
					tmpStr = Replace(tmpStr, "{transtype}", getcxcDataLngStr("LtxtTransTypeJE"))
				Case "57" 
					tmpStr = Replace(tmpStr, "{transtype}", getcxcDataLngStr("LtxtTransTypeCP"))
				Case "46" 
					tmpStr = Replace(tmpStr, "{transtype}", getcxcDataLngStr("LtxtTransTypePS"))
				Case "18" 
					tmpStr = Replace(tmpStr, "{transtype}", getcxcDataLngStr("LtxtTransTypePI"))
				Case "19" 
					tmpStr = Replace(tmpStr, "{transtype}", getcxcDataLngStr("LtxtTransTypeDN"))
				Case Else
					tmpStr = Replace(tmpStr, "{transtype}", "")
			End Select
			
			tmpStr = Replace(tmpStr, "{ref1}", rc("ref1"))
			tmpStr = Replace(tmpStr, "{linememo}", rc("linememo"))
			tmpStr = Replace(tmpStr, "{CreatedBy}", rc("CreatedBy"))

			If Rc("debit") <> "0" or rc("TransType") = "13" Then
				curDebit = CDbl(rc("debit"))
				If DirectRate = "Y" Then curDebit = curDebit/CurRate Else curDebit = curDebit*CurRate
				If Rc("debit") < "0" Then
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startIncTransDebitNeg", "endIncTransDebitNeg"), getMid(tmpStr, "startIncTransDebitNeg", "endIncTransDebitNeg"))
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startIncTransDebit", "endIncTransDebit"), "")
					tmpStr = Replace(tmpStr, "{curDebit}", rs("Currency") & "&nbsp;" & FormatNumber(curDebit*-1,myApp.SumDec))
				Else
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startIncTransDebitNeg", "endIncTransDebitNeg"), "")
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startIncTransDebit", "endIncTransDebit"), getMid(tmpStr, "startIncTransDebit", "endIncTransDebit"))
					tmpStr = Replace(tmpStr, "{curDebit}", rs("Currency") & "&nbsp;" & FormatNumber(curDebit,myApp.SumDec))
				End If
			Else
				tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startIncTransDebitNeg", "endIncTransDebit"), "")
			End If
			
			If Rc("credit") <> "0" Then
				curCredit = CDbl(rc("credit"))
				If DirectRate = "Y" Then curCredit = curCredit/CurRate Else curCredit = curCredit*CurRate
				If Rc("credit") < "0" Then
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startIncTransCreditNeg", "endIncTransCreditNeg"), getMid(tmpStr, "startIncTransCreditNeg", "endIncTransCreditNeg"))
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startIncTransCredit", "endIncTransCredit"), "")
					tmpStr = Replace(tmpStr, "{curCredit}", rs("Currency") & "&nbsp;" & FormatNumber(curCredit*-1,myApp.SumDec))
				Else
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startIncTransCreditNeg", "endIncTransCreditNeg"), "")
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startIncTransCredit", "endIncTransCredit"), getMid(tmpStr, "startIncTransCredit", "endIncTransCredit"))
					tmpStr = Replace(tmpStr, "{curCredit}", rs("Currency") & "&nbsp;" & FormatNumber(curCredit,myApp.SumDec))
				End If
			Else
				tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startIncTransCreditNeg", "endIncTransCredit"), "")
			End If
			
			varx = CDbl(rc("tempsaldo")) + varx
			If DirectRate = "Y" Then tempSaldo = CDbl(varx)/CurRate Else tempSaldo = CDbl(varx)*CurRate
			varx = varx * DispPosDeb
			If tempSaldo < 0 then
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startIncTransBalNeg", "endIncTransBalNeg"), getMid(tmpStr, "startIncTransBalNeg", "endIncTransBalNeg"))
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startIncTransBal", "endIncTransBal"), "")
					tmpStr = Replace(tmpStr, "{tempSaldo}", rs("Currency") & "&nbsp;" & FormatNumber(tempSaldo*-1,myApp.SumDec))
			Else
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startIncTransBalNeg", "endIncTransBalNeg"), "")
					tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startIncTransBal", "endIncTransBal"), getMid(tmpStr, "startIncTransBal", "endIncTransBal"))
					tmpStr = Replace(tmpStr, "{tempSaldo}", rs("Currency") & "&nbsp;" & FormatNumber(tempSaldo,myApp.SumDec))
			End If
			varx = varx * DispPosDeb
			
			strLoopEnd = strLoopEnd & tmpStr
		rc.movenext
		loop
		strContent = Replace(strContent, getFullMid(strContent, "startIncTransLoop", "endIncTransLoop"), strLoopEnd)
		
		If DirectRate = "Y" Then varx = varx/CurRate Else varx = varx*CurRate
		varx = varx * DispPosDeb
		If varx < 0 Then
			strContent = Replace(strContent, getFullMid(strContent, "startIncTransBalNeg", "endIncTransBalNeg"), getMid(strContent, "startIncTransBalNeg", "endIncTransBalNeg"))
			strContent = Replace(strContent, getFullMid(strContent, "startIncTransBal", "endIncTransBal"), "")
			strContent = Replace(strContent, "{varxIncTrans}", rs("currency") & "&nbsp;" & FormatNumber(varx*-1,myApp.SumDec))
		Else
			strContent = Replace(strContent, getFullMid(strContent, "startIncTransBalNeg", "endIncTransBalNeg"), "")
			strContent = Replace(strContent, getFullMid(strContent, "startIncTransBal", "endIncTransBal"), getMid(strContent, "startIncTransBal", "endIncTransBal"))
			strContent = Replace(strContent, "{varxIncTrans}", rs("currency") & "&nbsp;" & FormatNumber(varx,myApp.SumDec))
		End If
		varx = varx * DispPosDeb
	Else
		strContent = Replace(strContent, getFullMid(strContent, "startIncTrans", "endIncTrans"), "")
	End If
Else
	strContent = Replace(strContent, getFullMid(strContent, "startIncTrans", "endIncTrans"), "")
End If

If myApp.GetShowCxcOpenInv Then
	Curr = rs("Currency")
	sql = 	"select DocDate, DocDueDate, DocNum, T0.DocEntry, IsNull(PymntGroup, '') As Condicion, DocTotal, DocTotal-PaidToDate As Saldo, " & _
			"DateDiff(day," & myApp.GetShowCxcOpenInvBy & ",getdate()) As Antiguedad " & _
			"from oinv T0 inner join octg T1 on T1.groupnum = T0.groupnum " & _
			"where cardcode = N'" & saveHTMLDecode(Session("UserName"), False) & "' and DocStatus = 'O' and DocTotal-PaidToDate <> 0 " & _
			"order by Antiguedad desc"
	set rs = conn.execute(sql)
	If Not rs.Eof Then
		strContent = Replace(strContent, getFullMid(strContent, "startOpenTrans", "endOpenTrans"), getMid(strContent, "startOpenTrans", "endOpenTrans"))
		strLoopEnd = ""
		do while not rs.eof
			tmpStr = getMid(strContent, "startOpenTransLoop", "endOpenTransLoop")
			tmpStr = Replace(tmpStr, "{DocNum}", rs("DocNum"))
			tmpStr = Replace(tmpStr, "{DocDate}", FormatDate(rs("DocDate"), True))
			tmpStr = Replace(tmpStr, "{DocDueDate}", FormatDate(rs("DocDueDate"), True))
			tmpStr = Replace(tmpStr, "{Condicion}", Rs("Condicion"))
			tmpStr = Replace(tmpStr, "{DocTotal}", curr & " " & FormatNumber(Rs("DocTotal"),myApp.SumDec))
			tmpStr = Replace(tmpStr, "{Saldo}", curr & " " & FormatNumber(Rs("Saldo"),myApp.SumDec))
			If Request("excell") <> "Y" Then
				tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startOpenTransLink", "endOpenTransLink"), getMid(tmpStr, "startOpenTransLink", "endOpenTransLink"))
				tmpStr = Replace(tmpStr, "{DocEntry}", rs("DocEntry"))
			Else
				tmpStr = Replace(tmpStr, getFullMid(tmpStr, "startOpenTransLink", "endOpenTransLink"), "")
			End If
			If CInt(Rs("Antiguedad")) < 0 Then
				tmpStr = Replace(tmpStr, "{Antiguedad}", "0")
			Else
				tmpStr = Replace(tmpStr, "{Antiguedad}", rs("Antiguedad"))
			End If
			strLoopEnd = strLoopEnd & tmpStr
		rs.movenext
		loop
		strContent = Replace(strContent, getFullMid(strContent, "startOpenTransLoop", "endOpenTransLoop"), strLoopEnd)
	Else
		strContent = Replace(strContent, getFullMid(strContent, "startOpenTrans", "endOpenTrans"), "")
	End If
Else
	strContent = Replace(strContent, getFullMid(strContent, "startOpenTrans", "endOpenTrans"), "")
End If
%>
<form method="POST" action="cxc.asp" name="Form1"><div id="dvCXCPrint"><%=strContent%></div>
<input type="hidden" name="LinkRep" value="<%=Request("LinkRep")%>">
<input type="hidden" name="c1" value="<%=Server.HTMLEncode(Session("UserName"))%>">
</form>
       
<form target="_blank" method="post" name="frmViewDetail" action="">
<input type="hidden" name="DocEntry" value="">
<input type="hidden" name="DocType" value="">
<input type="hidden" name="CardCode" value="">
<input type="hidden" name="pop" value="Y">
</form>
<script language="javascript">
function goDetail(DocType, DocEntry) {
	if (DocType == 2)
	{
		document.frmViewDetail.action = 'addCard/crdConfDetailOpen.asp';
		document.frmViewDetail.CardCode.value = DocEntry;
	}
	else if (DocType != 24 && DocType != 46)
	{
		document.frmViewDetail.action = "cxcDocDetailOpen.asp";
		document.frmViewDetail.DocEntry.value = DocEntry;
	}
	else
	{
		document.frmViewDetail.action = "cxcRctDetailOpen.asp";
		document.frmViewDetail.DocEntry.value = DocEntry;
	}
	document.frmViewDetail.DocType.value = DocType;
	document.frmViewDetail.submit();
}
function doCxcExport(Type)
{
	switch (Type)
	{
		case 'pdf':
			document.frmCxcExport.action = 'cxcPDF.asp';
			document.frmCxcExport.pdf.value = 'Y';
			document.frmCxcExport.excell.value = '';
			break;
		case 'excell':
			document.frmCxcExport.action = 'cxcPrint.asp';
			document.frmCxcExport.pdf.value = '';
			document.frmCxcExport.excell.value = 'Y';
			break;
	}
	document.frmCxcExport.submit();
}
</script>
<form method="post" name="frmCxcExport" action="">
<input type="hidden" name="timestamp1" value="<%=vardate1%>">
<input type="hidden" name="timestamp2" value="<%=vardate2%>">
<input type="hidden" name="excell" value="">
<input type="hidden" name="pdf" value="">
</form>
<%
set rs = nothing
set rd = nothing
set rc = nothing
set rxy1 = nothing
If Request("excell") <> "Y" Then %>
<script type="text/javascript">
    Calendar.setup({
        inputField     :    "timestamp1",     // id of the input field
        ifFormat       :    "<%=GetCalendarFormatString%>",      // format of the input field
        button         :    "btntimestamp1",  // trigger for the calendar (button ID)
        align          :    "Bl",           // alignment (defaults to "Bl")
        singleClick    :    true
    });
    Calendar.setup({
        inputField     :    "timestamp2",     // id of the input field
        ifFormat       :    "<%=GetCalendarFormatString%>",      // format of the input field
        button         :    "btntimestamp2",  // trigger for the calendar (button ID)
        align          :    "Br",           // alignment (defaults to "Bl")
        singleClick    :    true
    });
</script>
<% End If %>
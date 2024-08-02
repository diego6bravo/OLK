<% addLngPathStr = "portal/" %>
<!--#include file="lang/openDocs.asp" -->
<%

ShowRep = False

If Request("order1") = "" Then order1 = "dType" Else order1 = Request("order1")
If Request("order2") = "" Then order2 = "desc" Else order2 = Request("order2")

SlpFilter = ""
If Not myAut.HasAuthorization(97) Then SlpFilter = " and T0.SlpCode = " & Session("vendid") & " "


varCardCode = saveHTMLDecode(Session("UserName"), False)

sql = ""

If CardType <> "S" Then
	ViewInvoice = True
	ViewOrder = True
	ViewQuote = True
	ViewDelivery = True
	
	If userType = "V" Then
		ViewInvoice = myAut.GetObjectProperty(13, "V")
		ViewOrder = myAut.GetObjectProperty(17, "V")
		ViewQuote = myAut.GetObjectProperty(23, "V")
		ViewDelivery = myAut.GetObjectProperty(15, "V")
	End If
	
	ShowRep = ViewInvoice or ViewOrder or ViewQuote or ViewDelivery 

	if ViewInvoice and (Request("DocType") = "" or Request("DocType") = "13" or Request("DocType") = "13R") Then 
		sql = sql & "select 13 dType, N'" & txtInv & "' + Case T0.IsIns When 'N' Then '' When 'Y' Then N' (" & getopenDocsLngStr("LtxtReserved") & ")' End DocType, T0.DocEntry, T0.DocNum, T0.DocCur, " & _
					"Case When T0.DocCur = N'" & myApp.MainCur & "' Then T0.DocTotal Else T0.DocTotalFC End DocTotal, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OSLP', 'SlpName', T0.SlpCode, SlpName) SlpName, T0.DocDate, T0.DocDueDate, Convert(int,T0.DocDate) DocDateInt, Convert(int,T0.DocDueDate) DueDateInt,  " & _
					"DateDiff(day,getdate(),T0.DocDate) Old, " & _
					"Case When T0.DocCur = N'" & myApp.MainCur & "' Then T0.DocTotal Else T0.DocTotalFC End - Case When T0.DocCur = N'" & myApp.MainCur & "' Then T0.PaidToDate Else T0.PaidFC End Saldo " & _
					"from OINV T0 " & _
					"inner join OSLP T1 on T1.SlpCode = T0.SlpCode " & _
					"where T0.CardCode = N'" & varCardCode & "' " & SlpFilter & " and ("
	
		If Request("DocType") = "" Then			
			sql = sql & "IsIns = 'N' and DocStatus = 'O' or IsIns = 'Y' and (DocStatus = 'O' or InvntSttus = 'O')"
		ElseIf Request("DocType") = "13" Then
			sql = sql & "IsIns = 'N' and DocStatus = 'O' or IsIns = 'Y' and (DocStatus = 'O')"
		ElseIf Request("DocType") = "13R" Then
			sql = sql & "IsIns = 'Y' and (InvntSttus = 'O')"
		End If 
		
		sql = sql & ") "
	End If
	
	If ViewOrder and (Request("DocType") = "" or Request("DocType") = "17") Then
		If sql <> "" Then sql = sql & "union "
		sql = sql & "select 17 dType, N'" & txtOrdr & "' DocType, T0.DocEntry, T0.DocNum, T0.DocCur, " & _
					"Case When T0.DocCur = N'" & myApp.MainCur & "' Then T0.DocTotal Else T0.DocTotalFC End DocTotal, " & _
					"OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OSLP', 'SlpName', T0.SlpCode, SlpName) SlpName, T0.DocDate, T0.DocDueDate, Convert(int,T0.DocDate) DocDateInt, Convert(int,T0.DocDueDate) DueDateInt,  " & _
					"DateDiff(day,getdate(),T0.DocDate) Old, " & _
					"Case When T0.DocCur = N'" & myApp.MainCur & "' Then T0.DocTotal Else T0.DocTotalFC End - Case When T0.DocCur = N'" & myApp.MainCur & "' Then T0.PaidToDate Else T0.PaidFC End Saldo " & _
					"from ORDR T0 " & _
					"inner join OSLP T1 on T1.SlpCode = T0.SlpCode " & _
					"where T0.CardCode = N'" & varCardCode & "' " & SlpFilter & " and DocStatus = 'O' "
	End If

	If ViewQuote and (Request("DocType") = "" or Request("DocType") = "23") Then
		If sql <> "" Then sql = sql & "union "
		sql = sql & "select 23 dType, N'" & txtQuote & "' DocType, T0.DocEntry, T0.DocNum, T0.DocCur, " & _
				"Case When T0.DocCur = N'" & myApp.MainCur & "' Then T0.DocTotal Else T0.DocTotalFC End DocTotal, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OSLP', 'SlpName', T0.SlpCode, SlpName) SlpName, T0.DocDate, T0.DocDueDate, Convert(int,T0.DocDate) DocDateInt, Convert(int,T0.DocDueDate) DueDateInt,  " & _
				"DateDiff(day,getdate(),T0.DocDate) Old, " & _
				"Case When T0.DocCur = N'" & myApp.MainCur & "' Then T0.DocTotal Else T0.DocTotalFC End - Case When T0.DocCur = N'" & myApp.MainCur & "' Then T0.PaidToDate Else T0.PaidFC End Saldo " & _
				"from OQUT T0 " & _
				"inner join OSLP T1 on T1.SlpCode = T0.SlpCode " & _
				"where T0.CardCode = N'" & varCardCode & "' " & SlpFilter & " and DocStatus = 'O' "
	End If
	
	If ViewDelivery and (Request("DocType") = "" or Request("DocType") = "15") Then
		If sql <> "" Then sql = sql & "union "
		sql = sql & "select 15 dType, N'" & txtOdln & "' DocType, T0.DocEntry, T0.DocNum, T0.DocCur, " & _
				"Case When T0.DocCur = N'" & myApp.MainCur & "' Then T0.DocTotal Else T0.DocTotalFC End DocTotal, T1.SlpName, T0.DocDate, T0.DocDueDate, Convert(int,T0.DocDate) DocDateInt, Convert(int,T0.DocDueDate) DueDateInt,  " & _
				"DateDiff(day,getdate(),T0.DocDate) Old, " & _
				"Case When T0.DocCur = N'" & myApp.MainCur & "' Then T0.DocTotal Else T0.DocTotalFC End - Case When T0.DocCur = N'" & myApp.MainCur & "' Then T0.PaidToDate Else T0.PaidFC End Saldo " & _
				"from ODLN T0 " & _
				"inner join OSLP T1 on T1.SlpCode = T0.SlpCode " & _
				"where T0.CardCode = N'" & varCardCode & "' " & SlpFilter & " and DocStatus = 'O' "
	End If
ElseIf CardType = "S" Then
	ShowRep = myAut.GetObjectProperty(18, "V") or myAut.GetObjectProperty(22, "V") or myAut.GetObjectProperty(20, "V")

	if myAut.GetObjectProperty(18, "V")  and (Request("DocType") = "" or Request("DocType") = "18" or Request("DocType") = "18R") Then 
		sql = sql & "select 18 dType, N'" & txtOpch & "' + Case T0.IsIns When 'N' Then '' When 'Y' Then N' (" & getopenDocsLngStr("LtxtReserved") & ")' End DocType, T0.DocEntry, T0.DocNum, T0.DocCur, " & _
					"Case When T0.DocCur = N'" & myApp.MainCur & "' Then T0.DocTotal Else T0.DocTotalFC End DocTotal, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OSLP', 'SlpName', T0.SlpCode, SlpName) SlpName, T0.DocDate, T0.DocDueDate, Convert(int,T0.DocDate) DocDateInt, Convert(int,T0.DocDueDate) DueDateInt,  " & _
					"DateDiff(day,getdate(),T0.DocDate) Old, " & _
					"Case When T0.DocCur = N'" & myApp.MainCur & "' Then T0.DocTotal Else T0.DocTotalFC End - Case When T0.DocCur = N'" & myApp.MainCur & "' Then T0.PaidToDate Else T0.PaidFC End Saldo " & _
					"from OPCH T0 " & _
					"inner join OSLP T1 on T1.SlpCode = T0.SlpCode " & _
					"where T0.CardCode = N'" & varCardCode & "' " & SlpFilter & " and ("
	
		If Request("DocType") = "" Then			
			sql = sql & "IsIns = 'N' and DocStatus = 'O' or IsIns = 'Y' and (DocStatus = 'O' or InvntSttus = 'O')"
		ElseIf Request("DocType") = "18" Then
			sql = sql & "IsIns = 'N' and DocStatus = 'O' or IsIns = 'Y' and (DocStatus = 'O')"
		ElseIf Request("DocType") = "18R" Then
			sql = sql & "IsIns = 'Y' and (InvntSttus = 'O')"
		End If 
		
		sql = sql & ") "
	End If 
	
	
	If myAut.GetObjectProperty(22, "V") and (Request("DocType") = "" or Request("DocType") = "22") Then
		If sql <> "" Then sql = sql & "union "
		sql = sql & "select 22 dType, N'" & txtOpor & "' DocType, T0.DocEntry, T0.DocNum, T0.DocCur, " & _
					"Case When T0.DocCur = N'" & myApp.MainCur & "' Then T0.DocTotal Else T0.DocTotalFC End DocTotal, " & _
					"OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OSLP', 'SlpName', T0.SlpCode, SlpName) SlpName, T0.DocDate, T0.DocDueDate, Convert(int,T0.DocDate) DocDateInt, Convert(int,T0.DocDueDate) DueDateInt,  " & _
					"DateDiff(day,getdate(),T0.DocDate) Old, " & _
					"Case When T0.DocCur = N'" & myApp.MainCur & "' Then T0.DocTotal Else T0.DocTotalFC End - Case When T0.DocCur = N'" & myApp.MainCur & "' Then T0.PaidToDate Else T0.PaidFC End Saldo " & _
					"from OPOR T0 " & _
					"inner join OSLP T1 on T1.SlpCode = T0.SlpCode " & _
					"where T0.CardCode = N'" & varCardCode & "' " & SlpFilter & " and DocStatus = 'O' "
	End If
	
	
	If myAut.GetObjectProperty(20, "V") and (Request("DocType") = "" or Request("DocType") = "20") Then
		If sql <> "" Then sql = sql & "union "
		sql = sql & "select 20 dType, N'" & txtOpdn & "' DocType, T0.DocEntry, T0.DocNum, T0.DocCur, " & _
				"Case When T0.DocCur = N'" & myApp.MainCur & "' Then T0.DocTotal Else T0.DocTotalFC End DocTotal, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OSLP', 'SlpName', T0.SlpCode, SlpName) SlpName, T0.DocDate, T0.DocDueDate, Convert(int,T0.DocDate) DocDateInt, Convert(int,T0.DocDueDate) DueDateInt,  " & _
				"DateDiff(day,getdate(),T0.DocDate) Old, " & _
				"Case When T0.DocCur = N'" & myApp.MainCur & "' Then T0.DocTotal Else T0.DocTotalFC End - Case When T0.DocCur = N'" & myApp.MainCur & "' Then T0.PaidToDate Else T0.PaidFC End Saldo " & _
				"from OPDN T0 " & _
				"inner join OSLP T1 on T1.SlpCode = T0.SlpCode " & _
				"where T0.CardCode = N'" & varCardCode & "' " & SlpFilter & " and DocStatus = 'O' "
	End If

End If

If ShowRep Then

	sql = "select * from (" & sql & ") As Table1 order by " & order1 & " " & order2
	
	If order1 = "dType" Then sql = sql & ", DueDateInt "
	set rs = Server.CreateObject("ADODB.RecordSet")
	rs.open sql, conn, 3, 1
	
	rs.PageSize = 40
	iPageCount = rs.PageCount

	If Request("Page") <> "" Then iPageCurrent = CLng(Request("Page")) Else iPageCurrent = 1
	
	iNextCount = iPageCurrent
	iCurMax = iPageCount/15
	iCurNext = 0
	do while iNextCount > 0
		iNextCount = iNextCount - 15
		iCurNext = iCurNext + 1
	loop
	If iCurMax - CInt(iCurMax) > 0 Then iCurMax = CInt(iCurMax) + 1
	fromI = (iCurNext*15)-14
	toI = (iCurNext*15)

	If iCurMax <= iCurNext Then toI = iPageCount
	If iPageCurrent > iPageCount Then iPageCurrent = iPageCount
	If iPageCurrent < 1 Then iPageCurrent = 1

	If Not rs.Eof Then rs.AbsolutePage = iPageCurrent 
End If

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetObjectData" & Session("ID")
cmd.Parameters.Refresh()
cmd("@ObjType") = "S"
cmd("@ObjID") = 14
cmd("@UserType") = userType
set ra = cmd.execute()
strContent = ra("ObjContent")
strContent = Replace(strContent, "{SelDes}", SelDes)
strContent = Replace(strContent, "{rtl}", Session("rtl"))
strContent = Replace(strContent, "{LttlPendLst}", getopenDocsLngStr("LttlPendLst"))
strContent = Replace(strContent, "{LtxtDocType}", getopenDocsLngStr("LtxtDocType"))
strContent = Replace(strContent, "{DtxtAll}", getopenDocsLngStr("DtxtAll"))

If InStr(strContent, "{DocTypeOptions}") = 0 Then
	strContent = Replace(strContent, "{txtQuotes}", txtQuotes)
	strContent = Replace(strContent, "{txtOrdrs}", txtOrdrs)
	strContent = Replace(strContent, "{txtOdlns}", txtOdlns)
	strContent = Replace(strContent, "{txtInvs}", txtInvs)
	
	If Request("DocType") = "23" Then strContent = Replace(strContent, "{SelDocType23}", "selected") Else strContent = Replace(strContent, "{SelDocType23}", "")
	If Request("DocType") = "17" Then strContent = Replace(strContent, "{SelDocType17}", "selected") Else strContent = Replace(strContent, "{SelDocType17}", "")
	If Request("DocType") = "15" Then strContent = Replace(strContent, "{SelDocType15}", "selected") Else strContent = Replace(strContent, "{SelDocType15}", "")
	If Request("DocType") = "13" Then strContent = Replace(strContent, "{SelDocType13}", "selected") Else strContent = Replace(strContent, "{SelDocType13}", "")
Else
	optStr = ""
	
	arrObjsStr = ""
	If userType = "V" Then
		If CardType = "C" or CardType = "L" Then
			If myAut.GetObjectProperty(23, "V") Then arrObjsStr = "23"
			If myAut.GetObjectProperty(17, "V") Then arrObjsStr = myAut.ConcValue(arrObjsStr, "17")
			
			If CardType = "C" Then
				If myAut.GetObjectProperty(15, "V") Then arrObjsStr = myAut.ConcValue(arrObjsStr, "15")
				If myAut.GetObjectProperty(13, "V") Then arrObjsStr = myAut.ConcValue(arrObjsStr, "13, 13R")
			End If
		ElseIf CardType = "S" Then
			If myAut.GetObjectProperty(22, "V") Then arrObjsStr = "22"
			If myAut.GetObjectProperty(20, "V") Then arrObjsStr = myAut.ConcValue(arrObjsStr, "20")
			If myAut.GetObjectProperty(18, "V") Then arrObjsStr = myAut.ConcValue(arrObjsStr, "18, 18R")
		End If
	Else
		arrObjsStr = "23, 17, 15, 13, 13R"
	End If
	
	If arrObjsStr <> "" Then
		arrObjs = Split(arrObjsStr, ", ")
		For i = 0 to UBound(arrObjs)
			objCode = arrObjs(i)
			
			optStr = optStr & "<option"
			
			If Request("DocType") = objCode Then optStr = optStr & " selected"
			
			optStr = optStr & " value = """ & objCode & """>"
			Select Case objCode
				Case "23"
					optStr = optStr & txtQuotes
				Case "17"
					optStr = optStr & txtOrdrs
				Case "15"
					optStr = optStr & txtOdlns
				Case "13"
					optStr = optStr & txtInvs
				Case "13R"
					optStr = optStr & txtInvs & " (" & getopenDocsLngStr("LtxtReserved") & ")"
				Case "22"
					optStr = optStr & txtOpors
				Case "18"
					optStr = optStr & txtOpchs
				Case "18R"
					optStr = optStr & txtOpchs & " (" & getopenDocsLngStr("LtxtReserved") & ")"
				Case "20"
					optStr = optStr & txtOpdns
			End Select
			optStr = optStr & "</option>"
		Next
		strContent = Replace(strContent, "{DocTypeOptions}", optStr)
	Else
		strContent = Replace(strContent, "{DocTypeOptions}", "")
	End If
End If

strContent = Replace(strContent, "{DtxtType}", getopenDocsLngStr("DtxtType"))
strContent = Replace(strContent, "{LtxtDocNum}", getopenDocsLngStr("LtxtDocNum"))
strContent = Replace(strContent, "{DtxtDate}", getopenDocsLngStr("DtxtDate"))
strContent = Replace(strContent, "{LtxtDue}", getopenDocsLngStr("LtxtDue"))
strContent = Replace(strContent, "{txtAgent}", txtAgent)
strContent = Replace(strContent, "{DtxtTotal}", getopenDocsLngStr("DtxtTotal"))
strContent = Replace(strContent, "{DtxtBalance}", getopenDocsLngStr("DtxtBalance"))
strContent = Replace(strContent, "{LtxtAntiq}", getopenDocsLngStr("LtxtAntiq"))
strContent = Replace(strContent, "{MainDoc}", MainDoc)
If Session("rtl") = "" Then
	strContent = Replace(strContent, "{rtl}", "")
Else
	strContent = Replace(strContent, "{rtl}", "rtl/")
End If

strContent = Replace(strContent, "{bgColDTypeImg}", doOpenDocsSortImg("dType"))
strContent = Replace(strContent, "{bgColDType}", doOpenDocsSortBG("dType"))

strContent = Replace(strContent, "{bgColDocNumImg}", doOpenDocsSortImg("DocNum"))
strContent = Replace(strContent, "{bgColDocNum}", doOpenDocsSortBG("DocNum"))

strContent = Replace(strContent, "{bgColDocDateIntImg}", doOpenDocsSortImg("DocDateInt"))
strContent = Replace(strContent, "{bgColDocDateInt}", doOpenDocsSortBG("DocDateInt"))

strContent = Replace(strContent, "{bgColDueDateIntImg}", doOpenDocsSortImg("DueDateInt"))
strContent = Replace(strContent, "{bgColDueDateInt}", doOpenDocsSortBG("DueDateInt"))

strContent = Replace(strContent, "{bgColSlpNameImg}", doOpenDocsSortImg("SlpName"))
strContent = Replace(strContent, "{bgColSlpName}", doOpenDocsSortBG("SlpName"))

strContent = Replace(strContent, "{bgColDocTotalImg}", doOpenDocsSortImg("DocTotal"))
strContent = Replace(strContent, "{bgColDocTotal}", doOpenDocsSortBG("DocTotal"))

strContent = Replace(strContent, "{bgColSaldoImg}", doOpenDocsSortImg("Saldo"))
strContent = Replace(strContent, "{bgColSaldo}", doOpenDocsSortBG("Saldo"))

strContent = Replace(strContent, "{bgColOldImg}", doOpenDocsSortImg("Old"))
strContent = Replace(strContent, "{bgColOld}", doOpenDocsSortBG("Old"))

If ShowRep and Not rs.Eof Then
	strContent = Replace(strContent, getFullMid(strContent, "startNoData", "endNoData"), "")
	
	strLoop = ""
		for intRecord=1 to rs.PageSize
		tmpStr = getMid(strContent, "startLoop", "endLoop")
		tmpStr = Replace(tmpStr, "{DocType}", rs("DocType"))
		tmpStr = Replace(tmpStr, "{DocEntry}", rs("DocEntry"))
		tmpStr = Replace(tmpStr, "{dType}", rs("dType"))
		tmpStr = Replace(tmpStr, "{DocNum}", rs("DocNum"))
		tmpStr = Replace(tmpStr, "{DocDate}", ConvDate(rs("DocDate"), "%d/%M/%y"))
		tmpStr = Replace(tmpStr, "{DocDueDate}", ConvDate(rs("DocDueDate"), "%d/%M/%y"))
		tmpStr = Replace(tmpStr, "{SlpName}", rs("SlpName"))
		tmpStr = Replace(tmpStr, "{DocTotal}", rs("DocCur") & " " & FormatNumber(rs("DocTotal"),myApp.SumDec))
		tmpStr = Replace(tmpStr, "{Saldo}", rs("DocCur") & " " & FormatNumber(rs("Saldo"),myApp.SumDec))
		tmpStr = Replace(tmpStr, "{Old}", "<span dir=""ltr"">" & rs("Old") & "</span>")
		
		strLoop = strLoop & tmpStr

		rs.movenext
		If rs.eof then exit for
	next
	strContent = Replace(strContent, getFullMid(strContent, "startLoop", "endLoop"), strLoop)
	
Else
	strContent = Replace(strContent, getFullMid(strContent, "startLoop", "endLoop"), "")
	strContent = Replace(strContent, getFullMid(strContent, "startNoData", "endNoData"), getMid(strContent, "startNoData", "endNoData"))
	strContent = Replace(strContent, "{DtxtNoData}", getopenDocsLngStr("DtxtNoData"))
End If


Response.Write Left(strContent, InStr(strContent, "<!--startPaging-->")-1)

Response.Write doPagingStr(getMid(strContent, "startPaging", "endPaging"))

Response.Write getMid(strContent, "endPaging", "startEndTable")

Response.Write doPagingStr(getMid(strContent, "startPaging", "endPaging"))

Response.Write getMid(strContent, "startEndTable", "endEndTable")

%>
<script language="javascript">
function goViewDoc(DocEntry, DocType)
{
	document.frmViewDet.DocEntry.value = DocEntry;
	document.frmViewDet.doctype.value = DocType;
	document.frmViewDet.submit();
}
function doSort(c)
{
	document.frmGoPage.order1.value = c;
	if ('<%=order1%>' == c)
	{
		if ('<%=order2%>' == 'asc')
			document.frmGoPage.order2.value = 'desc';
		else
			document.frmGoPage.order2.value = 'asc';
	}
	else
	{
		document.frmGoPage.order2.value = 'asc';
	}
	document.frmGoPage.submit();
}
function goType(t)
{
	document.frmGoPage.DocType.value = t;
	document.frmGoPage.submit();
}
function goPage(i) 
{
	document.frmGoPage.Page.value = i;
	document.frmGoPage.submit();
}
</script>
<form name="frmGoPage" action="<%=strScriptName%>" method="post">
<% For each itm in Request.Form %>
<input type="hidden" name="<%=itm%>" value="<%=Request(itm)%>">
<% Next %>
<% For each itm in Request.QueryString %>
<input type="hidden" name="<%=itm%>" value="<%=Request(itm)%>">
<% Next %>
<% If Request.Form.Count = 0 or strScriptName = "activeclient.asp" and Request("activeClient") <> "Y" Then %>
<input type="hidden" name="order1" value="">
<input type="hidden" name="order2" value="">
<input type="hidden" name="DocType" value="<%=Request("DocType")%>">
<input type="hidden" name="Page" value="<%=iPageCurrent%>">
<% If strScriptName = "activeclient.asp" Then %><input type="hidden" name="activeClient" value="Y"><% End If %>
<% End If %>
</form>
<form target="_blank" method="post" name="frmViewDet" action="cxcDocDetailOpen.asp">
<input type="hidden" name="DocEntry" value="">
<input type="hidden" name="doctype" value="">
<input type="hidden" name="pop" value="Y">
</form>

<%
Function doOpenDocsSortImg(c)
	If LCase(order1) = LCase(c) Then
		If order2 = "asc" Then
			doOpenDocsSortImg = "<img src=""images/arrow_up.gif"">"
		Else
			doOpenDocsSortImg = "<img src=""images/arrow_down.gif"">"
		End If
	Else
		doOpenDocsSortImg = ""
	End If
End Function
Function doOpenDocsSortBG(c)
	If LCase(order1) = LCase(c) Then doOpenDocsSortBG = "class=""GeneralTblBold2HighLight""" Else doOpenDocsSortBG = ""
End Function

%>
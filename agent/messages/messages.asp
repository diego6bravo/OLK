<% addLngPathStr = "messages/" %>
<!--#include file="lang/messages.asp" -->
<%
Select Case SelDes
	Case "1"
		bgColorIn1 = "#FBFDD7"
		bgColorIn2 = "#C9D3DC"
	Case "2"
		bgColorIn1 = "#FBFDD7"
		bgColorIn2 = "#FBFDDB"
	Case "3"
		bgColorIn1 = "#FBFDD7"
		bgColorIn2 = "#C5DFFC"
	Case Else
End Select
If Session("UserName") <> "-Anon-" Then %>
<SCRIPT LANGUAGE="JavaScript">
function Start(page) {
OpenWin = this.open(page, "messages", "toolbar=no,menubar=no,location=no,scrollbars=no,resizable=yes, width=400,height=360");
OpenWin.focus()
}
</SCRIPT>
<%
Select Case userType 
	Case "V" 
		user = Session("vendid")
		MainDoc = "agent"
	Case "C" 
		user = Session("UserName")
		MainDoc = "messages"
End Select 

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetMessages" & Session("ID")
cmd.Parameters.Refresh()
cmd("@UserType") = userType
cmd("@User") = user
If Request("msgOrdr1") <> "" Then cmd("@Order1") = Request("msgOrdr1")
If Request("msgOrdr2") <> "" Then cmd("@Order2") = Request("msgOrdr2")
set rs = Server.CreateObject("ADODB.recordset")
rs.open cmd, , 3, 1

If Request("msgOrdr1") = "" or Request("msgOrdr1") = "DateSort" Then 
	order1 = "T0.OlkDate" 
ElseIf Request("msgOrdr1") = "OlkType" Then
	order1 = "T0.OlkUFromType"
	If Request("msgOrdr2") <> "" Then order1 = order1 & " " & Request("msgOrdr2")
	order1 = order1 & ", T2.CardType"
	altOrder = "OlkType"
Else 
	order1 = Request("msgOrdr1")
End If
If Request("msgOrdr2") = "" Then order2 = "desc" Else order2 = Request("msgOrdr2")

rs.PageSize = 10
rs.CacheSize = 10

iPageCount = RS.PageCount

If Request("p") <> "" Then iCurPage = CInt(Request("p")) Else iCurPage = 1

iNextCount = iCurPage
iCurMax = iPageCount/15
iCurNext = 0
do while iNextCount > 0
	iNextCount = iNextCount - 15
	iCurNext = iCurNext + 1
loop
If iCurNext - CInt(iCurMax) > 0 Then iCurMax = CInt(iCurMax)+1

fromI = (iCurNext*15)-14
toI = (iCurNext*15)

If iCurPage > iPageCount Then iCurPage = iPageCount
If iCurPage < 1 Then iCurPage = 1
If toI > iPageCount Then toI = iPageCount

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetObjectData" & Session("ID")
cmd.Parameters.Refresh()
cmd("@ObjType") = "S"
cmd("@ObjID") = 13
cmd("@UserType") = userType
set ra = cmd.execute()
strContent = ra("ObjContent")
strContent = Replace(strContent, "{SelDes}", SelDes)
strContent = Replace(strContent, "{rtl}", Session("rtl"))
If Session("rtl") = "" Then
	strContent = Replace(strContent, "{rtlAlign}", "right")
Else
	strContent = Replace(strContent, "{rtlAlign}", "left")
End If
strContent = Replace(strContent, "{LttlMsgBox}", getmessagesLngStr("LttlMsgBox"))
strContent = Replace(strContent, "{DtxtDate}", getmessagesLngStr("DtxtDate"))
strContent = Replace(strContent, "{DtxtSubject}", getmessagesLngStr("DtxtSubject"))
strContent = Replace(strContent, "{LtxtDel}", getmessagesLngStr("LtxtDel"))
strContent = Replace(strContent, "{LtxtNewMsg}", getmessagesLngStr("LtxtNewMsg"))
strContent = Replace(strContent, "{LtxtDelete}", getmessagesLngStr("LtxtDelete"))
strContent = Replace(strContent, "{LtxtConfDelMsgs}", getmessagesLngStr("LtxtConfDelMsgs"))
strContent = Replace(strContent, "{redir}", Request("cmd"))
strContent = Replace(strContent, "{p}", iCurPage)
strContent = Replace(strContent, "{msgOrdr1}", Request("msgOrdr1"))
strContent = Replace(strContent, "{msgOrdr2}", Request("msgOrdr2"))
strContent = Replace(strContent, "{onlyMsg}", Request("onlyMsg"))

strContent = Replace(strContent, "{bgColStatusImg}", doMsgInboxSortImg("OlkStatus"))
strContent = Replace(strContent, "{bgColStatus}", doMsgInboxSortBG("OlkStatus"))

strContent = Replace(strContent, "{bgColUrgentImg}", doMsgInboxSortImg("OlkUrgent"))
strContent = Replace(strContent, "{bgColUrgent}", doMsgInboxSortBG("OlkUrgent"))

strContent = Replace(strContent, "{bgColDateImg}", doMsgInboxSortImg("T0.OlkDate"))
strContent = Replace(strContent, "{bgColDate}", doMsgInboxSortBG("T0.OlkDate"))

strContent = Replace(strContent, "{bgColSubjectImg}", doMsgInboxSortImg("OlkSubject"))
strContent = Replace(strContent, "{bgColSubject}", doMsgInboxSortBG("OlkSubject"))

strContent = Replace(strContent, "{bgColTypeImg}", doMsgInboxSortImg("OlkType"))
strContent = Replace(strContent, "{bgColType}", doMsgInboxSortBG("OlkType"))

If Not rs.Eof then
	strContent = Replace(strContent, getFullMid(strContent, "startNoData", "endNoData"), "")
	RS.AbsolutePage = iCurPage 
	strLoop = ""
	for intRecord=1 to rs.PageSize
		tmpStr = getMid(strContent, "startLoop", "endLoop")
		tmpStr = Replace(tmpStr, "{OlkLog}", rs("OlkLog"))
		tmpStr = Replace(tmpStr, "{Date}", FormatDate(RS("Date"), True))
		tmpStr = Replace(tmpStr, "{OlkSubject}", rs("OlkSubject"))
		If CStr(Trim(RS("OlkStatus"))) = "N" Then 
			OImg = "new" 
			OImgLnk = "O"
		Else 
			OImg = "open"
			OImgLnk = "N"
		End If
		tmpStr = Replace(tmpStr, "{OImg}", OImg)
		tmpStr = Replace(tmpStr, "{OImgLnk}", OImgLnk)
		
		If rs("OlkUrgent") = "Y" Then
			tmpStr = Replace(tmpStr, getFullMid(strContent, "startUrgent", "endUrgent"), getMid(strContent, "startUrgent", "endUrgent")) 
		Else
			tmpStr = Replace(tmpStr, getFullMid(strContent, "startUrgent", "endUrgent"), "") 
		End If
		
		imgStr = ""
		altStr = ""
		Select Case rs("OLKUFromType")
			Case "S"
				imgStr = "alert"
				altStr = getmessagesLngStr("DtxtAlert")
			Case "C"
				Select Case rs("CardType")
					Case "C"
						imgStr = "supplier"
						altStr = txtClient
					Case "S"
						imgStr = "client"
						altStr = getmessagesLngStr("DtxtSupplier")
					Case "L"
						imgStr = "lead"
						altStr = getmessagesLngStr("DtxtLead")
				End Select
			Case "V"
				imgStr = "agent"
				altStr = txtAgent
			Case "B"
				imgStr = "system"
				altStr = getmessagesLngStr("DtxtSystem")
			Case "E"
				imgStr = "alert_red"
				altStr = getmessagesLngStr("DtxtError")
		End Select
		tmpStr = Replace(tmpStr, "{TypeImg}", imgStr)
		tmpStr = Replace(tmpStr, "{TypeAlt}", altStr)

		strLoop = strLoop & tmpStr
	rs.movenext
	if rs.EOF then exit for
	next
	strContent = Replace(strContent, getFullMid(strContent, "startLoop", "endLoop"), strLoop)

	If iCurNext > 1 Then
		strPrevAll = Replace(getMid(strContent, "startPrevAll", "endPrevAll"), "{iLink}", ((iCurNext-1)*15))
	Else
		strPrevAll = ""
	End If

	If iCurPage > 1 Then
		strPrev = Replace(getMid(strContent, "startPrev", "endPrev"), "{iLink}", (iCurPage-1))
	Else
		strPrev = ""
	End If
	
	strLoop = ""
	strPagLink = getMid(strContent, "startPageLink", "endPageLink")
	strPagCur = getMid(strContent, "startPageCur", "endPageCur")
	If iPageCount > 1 Then
		For i = fromI to toI
			If i <> iCurPage Then
				strLoop = strLoop & Replace(strPagLink, "{iLink}", i)
			Else
				strLoop = strLoop & Replace(strPagCur, "{iLink}", i)
			End If
		Next
	End If
	strContent = Replace(strContent, getFullMid(strContent, "startPageLink", "endPageCur"), strLoop)
	
	If iCurPage < iPageCount Then
		strNext = Replace(getMid(strContent, "startNext", "endNext"), "{iLink}", (iCurPage+1))
	Else
		strNext = ""
	End If
	
	If iCurNext < iCurMax Then
		strNextAll = Replace(getMid(strContent, "startNextAll", "endNextAll"), "{iLink}", (iCurNext*15)+1)
	Else
		strNextAll = ""
	End If
	
	If Session("rtl") = "" Then
		strContent = Replace(strContent, getFullMid(strContent, "startPrev", "endPrev"), strPrev)
		strContent = Replace(strContent, getFullMid(strContent, "startPrevAll", "endPrevAll"), strPrevAll)
		strContent = Replace(strContent, getFullMid(strContent, "startNext", "endNext"), strNext)
		strContent = Replace(strContent, getFullMid(strContent, "startNextAll", "endNextAll"), strNextAll)
	Else
		strContent = Replace(strContent, getFullMid(strContent, "startPrev", "endPrev"), strNext)
		strContent = Replace(strContent, getFullMid(strContent, "startPrevAll", "endPrevAll"), strNextAll)
		strContent = Replace(strContent, getFullMid(strContent, "startNext", "endNext"), strPrev)
		strContent = Replace(strContent, getFullMid(strContent, "startNextAll", "endNextAll"), strPrevAll)
	End If
		
Else
	strContent = Replace(strContent, getFullMid(strContent, "startLoop", "endPaging"), "")
	strContent = Replace(strContent, getFullMid(strContent, "startNoData", "endNoData"), getMid(strContent, "startNoData", "endNoData"))
	strContent = Replace(strContent, "{LtxtNoMsg}", getmessagesLngStr("LtxtNoMsg"))
End If

If userType = "V" Then
	strContent = Replace(strContent, getFullMid(strContent, "startClientNewMsg", "endClientNewMsg"), "")
	strContent = Replace(strContent, getFullMid(strContent, "startAgentNewMsg", "endAgentNewMsg"), getMid(strContent, "startAgentNewMsg", "endAgentNewMsg"))
Else
	strContent = Replace(strContent, getFullMid(strContent, "startClientNewMsg", "endClientNewMsg"), getMid(strContent, "startClientNewMsg", "endClientNewMsg"))
	strContent = Replace(strContent, getFullMid(strContent, "startAgentNewMsg", "endAgentNewMsg"), "")
End If
strContent = Replace(strContent, "{cmd}", Request("cmd"))
%>

<%=strContent%>

<% End If %>
<script language="javascript">
function viewMsg(OlkLog)
{
	window.location.href = 'messagedetail.asp?olklog=' + OlkLog;
}
function delMsg(OlkLog)
{
	if(confirm('<%=getmessagesLngStr("LtxtConfDelMsg")%>'))window.location.href='messages/msgdelete.asp?olklog=' + OlkLog + '&redir=<%=Request("cmd")%>&AddPath=../';
}
function changeMsgStatus(OlkLog, OImgLnk)
{
	window.location.href='messages/updateMsgStatus.asp?olklog=' + OlkLog + '&status=' + OImgLnk + '&AddPath=../&onlyMsg=<%=Request("onlyMsg")%>';
}
</script>
<form method="post" action="<%=MainDoc%>.asp" name="frmGoPage">
<% For each itm in Request.Form
If itm <> "p" Then %>
<input type="hidden" name="<%=itm%>" value="<%=Request(itm)%>">
<% End If
Next %>
<% For each itm in Request.QueryString
If itm <> "p" Then %>
<input type="hidden" name="<%=itm%>" value="<%=Request(itm)%>">
<% End If
Next %>
<input type="hidden" name="p" value="">
<% If Request.Form.Count = 0 Then %>
<input type="hidden" name="msgOrdr1" value="">
<input type="hidden" name="msgOrdr2" value="">
<% End If %>
</form>
<script language="javascript">
function doSort(c)
{
	document.frmGoPage.msgOrdr1.value = c;
	if ('<%=order1%>' == c || '<%=altOrder%>' == c)
	{
		if ('<%=order2%>' == 'asc')
			document.frmGoPage.msgOrdr2.value = 'desc';
		else
			document.frmGoPage.msgOrdr2.value = 'asc';
	}
	else
	{
		document.frmGoPage.msgOrdr2.value = 'asc';
	}
	document.frmGoPage.p.value = 1;
	document.frmGoPage.submit();
}
function goPage(p)
{
	document.frmGoPage.p.value = p;
	document.frmGoPage.submit();
}
</script>
<%
Function doMsgInboxSortImg(c)
	If LCase(order1) = LCase(c) or altOrder = c Then
		If order2 = "asc" Then
			If alterSortImgUp = "" Then
				doMsgInboxSortImg = "<img src=""images/arrow_up.gif"">"
			Else
				doMsgInboxSortImg = alterSortImgUp
			End If
		Else
			If alterSortImgDown = "" Then
				doMsgInboxSortImg = "<img src=""images/arrow_down.gif"">"
			Else
				doMsgInboxSortImg = alterSortImgDown
			End If
		End If
	Else
		doMsgInboxSortImg = ""
	End If
End Function
Function doMsgInboxSortBG(c)
	If LCase(order1) = LCase(c) or altOrder = c Then doMsgInboxSortBG = "class=""GeneralTblBold2HighLight""" Else doMsgInboxSortBG = ""
End Function
%>
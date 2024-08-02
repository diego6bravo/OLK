<!--#include file="clientInc.asp"-->
<!--#include file="agentTop.asp"-->
<% If userType <> "V" Then Response.Redirect "default.asp" %>
<% If Not IsNumeric(Request("ADPollID")) Then Response.Redirect "unauthorized.asp" %>
<head>
<style type="text/css">
.style1 {
	text-align: center;
}
</style>
</head>

<% addLngPathStr = "" %>
<!--#include file="lang/extPollOpen.asp" -->
<% 
ADPollID = CInt(Request("ADPollID"))

sql = "select IsNull(T1.AlterName, T0.Name) Name, IsNull(T1.AlterDescription, T0.Description) Description, StartDate, EndDate, Filter " & _
"from OLKADPoll T0 " & _
"left outer join OLKADPollAlterNames T1 on T1.LanID = " & Session("LanID") & " and T1.AdPollID = T0.AdPollID " & _
"where T0.ADPollID = " & ADPollID
set rs = conn.execute(sql)
Name = rs("Name")
Description = rs("Description")
StartDate = FormatDate(rs("StartDate"), True)
EndDate = FormatDate(rs("EndDate"), True)
If rs("Filter") <> "" Then
	qcFilter = " and " & rs("Filter")
Else
	qcFilter = ""
End If

sql = "select DocEntry from OCRD where CardType in ('C', 'L') " & qcFilter & _
" and not exists(select '' from OLKADPollAnswers where ADPollID = " & ADPollID & " and CardCode = OCRD.CardCode) " & _
"order by CardCode"

set rs = Server.CreateObject("ADODB.RecordSet")
set rp = Server.CreateObject("ADODB.RecordSet")
rp.CursorLocation = 3
rp.open sql, conn, 3, 1

rp.PageSize = 40
nPageCount = rp.PageCount
If Request("Page") <> "" Then nPage = CLng(Request("Page")) Else nPage = 1

If Not rp.Eof Then
	rp.AbsolutePage = nPage

	DocEntry = ""
	do while not (rp.eof Or rp.AbsolutePage <> nPage)
		If DocEntry <> "" Then DocEntry = DocEntry & ", "
		DocEntry = DocEntry & rp("DocEntry")
	rp.MoveNext
	loop
	
	sql = "select CardCode, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRD', 'CardName', CardCode, CardName) CardName, Phone1, Phone2, CntctPrsn from OCRD where DocEntry in (" & DocEntry & ") " & _
	"order by CardCode"
	
	set rs = conn.execute(sql)
End If

iNextCount = nPage
iCurMax = nPageCount/15
iCurNext = 0
do while iNextCount > 0
iNextCount = iNextCount - 15
iCurNext = iCurNext + 1
loop
If iCurMax - CInt(iCurMax) > 0 Then iCurMax = CInt(iCurMax) + 1

fromI = (iCurNext*15)-14
toI = (iCurNext*15)

If iCurMax <= iCurNext Then toI = nPageCount
If nPage > nPageCount Then nPage = nPageCount
If nPage < 1 Then nPage = 1

%>
<table border="0" cellpadding="0" width="100%">
	<tr class="GeneralTlt">
		<td><%=getextPollOpenLngStr("LtxtExtPolls")%></td>
	</tr>
</table>
<table border="0" cellpadding="0" width="100%">
	<tr>
		<td width="120" class="GeneralTblBold2"><%=getextPollOpenLngStr("DtxtName")%></td>
		<td colspan="2" class="GeneralTbl"><%=Name%></td>
	</tr>
	<tr>
		<td width="120" class="GeneralTblBold2"><%=getextPollOpenLngStr("LtxtStartDate")%></td>
		<td width="250" class="GeneralTbl"><%=StartDate%></td>
		<td class="GeneralTblBold2"><%=getextPollOpenLngStr("DtxtDescription")%></td>
	</tr>
	<tr>
		<td width="120" class="GeneralTblBold2"><%=getextPollOpenLngStr("LtxtEndDate")%></td>
		<td width="250" class="GeneralTbl"><%=EndDate%></td>
		<td rowspan="3" valign="top" class="GeneralTbl"><% If Not IsNull(Description) Then %><%=Replace(Description,VbNewLine,"<br>")%><% End If %></td>
	</tr>
	<tr>
		<td width="120" class="GeneralTblBold2"><%=getextPollOpenLngStr("LtxtPending")%></td>
		<td width="250" class="GeneralTbl"><%=rp.recordcount%></td>
	</tr>
</table>
<table border="0" cellpadding="0" width="100%">
	<% If nPageCount > 1 Then %>
	<tr class="SearchPage">
		<td colspan="6">
		<div align="center">
			<table border="0" cellpadding="0" width="400" cellspacing="1" id="table4">
				<tr class="SearchPage">
					<td width="14">
					<% If iCurNext > 1 Then %><a href="javascript:goPage(<%= ((iCurNext-1)*15) %>);">
					<img border="0" src="design/0/images/<%=Session("rtl")%>prevAll.gif" width="12" height="13" align="left"></a><% End If %>
					</td>
					<td width="14">
					<% If nPage > 1 Then	%><a href="javascript:goPage(<%= nPage - 1 %>);">
					<img border="0" src="design/0/images/<%=Session("rtl")%>prev.gif" width="12" height="13" align="left"></a><% End If %></td>
					<td dir="ltr">
					<p align="center">
					<% 
					If nPageCount > 1 then
						For I = fromI To toI
							If I = nPage Then %>
							<font size="3">
							<b><%= I %></b></font>
							<% Else %>
							<a class="LnkSearchPaginacion" href="javascript:goPage(<%= I %>);"><%= I %></a>
							<% End If
						Next
					end if %></td>
					<td width="14">
					<% If nPage < nPageCount Then %>
		  <a href="javascript:goPage(<%= nPage + 1 %>);">
		<img border="0" src="design/<%=SelDes%>/images/<%=Session("rtl")%>next.gif" width="12" height="13" align="right">
		</a><% End If %>
		</td>
		<td width="14"><% If iCurNext < iCurMax Then %>
		  <a href="javascript:goPage(<%= (iCurNext*15)+1 %>);">
		  <img border="0" src="design/<%=SelDes%>/images/<%=Session("rtl")%>nextAll.gif" width="12" height="13" align="right"></a><% End If %></td>
				</tr>
			</table>
		</div>
		</td>
	</tr>
	<% End If %>
	<tr class="GeneralTblBold2">
		<td width="20">&nbsp;</td>
		<td class="style1"><%=getextPollOpenLngStr("DtxtCode")%></td>
		<td class="style1"><%=getextPollOpenLngStr("DtxtName")%></td>
		<td class="style1"><%=getextPollOpenLngStr("DtxtPhone")%>&nbsp;1</td>
		<td class="style1"><%=getextPollOpenLngStr("DtxtPhone")%>&nbsp;2</td>
		<td class="style1"><%=getextPollOpenLngStr("DtxtContact")%></td>
	</tr>
	<% If Rp.RecordCount > 0 Then %>
	<% do while not rs.eof %>
	<tr class="GeneralTbl">
		<td width="20">
		<a href="javascript:doMyLink('extPollExec.asp', 'AdPollID=<%=ADPollID%>&CardCode=<%=Replace(myHTMLEncode(rs("CardCode")), "'", "\'")%>', '_self');">
		<img border="0" src="images/<%=Session("rtl")%>flechaselec.gif" width="15" height="13"></a></td>
		<td><%=rs("CardCode")%></td>
		<td><%=rs("CardName")%></td>
		<td><%=rs("Phone1")%></td>
		<td><%=rs("Phone2")%></td>
		<td><%=rs("CntctPrsn")%></td>
	</tr>
	<% rs.movenext
	loop %>
	<% If nPageCount > 1 Then %>
	<tr class="SearchPage">
		<td colspan="6">
		<div align="center">
			<table border="0" cellpadding="0" width="400" cellspacing="1" id="table4">
				<tr class="SearchPage">
					<td width="14">
					<% If iCurNext > 1 Then %><a href="javascript:goPage(<%= ((iCurNext-1)*15) %>);">
					<img border="0" src="design/0/images/<%=Session("rtl")%>prevAll.gif" width="12" height="13" align="left"></a><% End If %>
					</td>
					<td width="14">
					<% If nPage > 1 Then	%><a href="javascript:goPage(<%= nPage - 1 %>);">
					<img border="0" src="design/0/images/<%=Session("rtl")%>prev.gif" width="12" height="13" align="left"></a><% End If %></td>
					<td>
					<p align="center">
					<% 
					If nPageCount > 1 then
						For I = fromI To toI
							If I = nPage Then %>
							<font size="3">
							<b><%= I %></b></font>
							<% Else %>
							<a class="LnkSearchPaginacion" href="javascript:goPage(<%= I %>);"><%= I %></a>
							<% End If
						Next
					end if %></td>
					<td width="14">
					<% If nPage < nPageCount Then %>
		  <a href="javascript:goPage(<%= nPage + 1 %>);">
		<img border="0" src="design/<%=SelDes%>/images/<%=Session("rtl")%>next.gif" width="12" height="13" align="right">
		</a><% End If %>
		</td>
		<td width="14"><% If iCurNext < iCurMax Then %>
		  <a href="javascript:goPage(<%= (iCurNext*15)+1 %>);">
		  <img border="0" src="design/<%=SelDes%>/images/<%=Session("rtl")%>nextAll.gif" width="12" height="13" align="right"></a><% End If %></td>
				</tr>
			</table>
		</div>
		</td>
	</tr>
	<% End If %>
	<% Else %>
	<tr>
		<td colspan="6" align="center" class="GeneralTbl"><%=getextPollOpenLngStr("DtxtNoData")%></td>
	</tr>
	<% End If %>
</table>
<form name="frmGPage" action="extPollOpen.asp" method="post">
<% For each itm in Request.Form
If itm <> "Page" and itm <> "submit" Then %>
<input type="hidden" name="<%=itm%>" value="<%=Request(itm)%>">
<% End If
Next %>
<% For each itm in Request.QueryString
If itm <> "Page" and itm <> "submit" Then %>
<input type="hidden" name="<%=itm%>" value="<%=Request(itm)%>">
<% End If
Next %>
<input type="hidden" name="Page" value="">
</form>
<script language="javascript">
function goPage(p) { document.frmGPage.Page.value = p; document.frmGPage.submit(); }
</script>

<!--#include file="agentBottom.asp"-->
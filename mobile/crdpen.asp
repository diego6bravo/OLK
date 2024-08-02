<!--#include file="lang/crdpen.asp" -->
<%
If Request("delLog") <> "" Then
	sql = "update R3_ObsCommon..TLOG set Status = 'B' where LogNum = " & Request("delLog")
	conn.execute(sql)
End If

sql = 	"select T0.LogNum " & _
		"from r3_obscommon..tlog T0 " & _
		"inner join r3_obscommon..TCRD T1 on T1.LogNum = T0.LogNum " & _
		"inner join R3_ObsCommon..TLOGControl L0 on L0.LogNum = T0.LogNum and L0.AppID = 'TM-OLK'  "
		
If Request("GroupNameFrom") <> "" or Request("GroupNameTo") <> "" Then
	sql = sql & "inner join OCRG T2 on T2.GroupCode = T1.GroupCode "
End If

If Request("CountryFrom") <> "" or Request("CountryTo") <> "" Then
	sql = sql & "inner join OCRY T3 on T3.Code = T1.Country collate database_default "
End If
		
sql = sql & "where T0.Object = 2 and T0.Status in ('R', 'H') and Company = '" & Session("olkDB") & "' "

If Request("LogNumFrom") <> "" Then sql = sql & " and T0.LogNum >= " & Request("LogNumFrom") & " "
If Request("LogNumTo") <> "" Then sql = sql & " and T0.LogNum <= " & Request("LogNumTo") & " "
If Request("dtFrom") <> "" Then sql = sql & " and DateDiff(day,Convert(datetime,'" & SaveSqlDate(Request("dtFrom")) & "',120), T0.SubDate) >= 0 "
If Request("dtTo") <> "" Then sql = sql & " and DateDiff(day,T0.SubDate, Convert(datetime,'" & SaveSqlDate(Request("dtTo")) & "',120)) >= 0"
If Request("CardCodeFrom") <> "" Then sql = sql & " and T1.CardCode >= N'" & saveHTMLDecode(Request("CardCodeFrom"), False) & "' "
If Request("CardCodeTo") <> "" Then sql = sql & " and T1.CardCode <= N'" & saveHTMLDecode(Request("CardCodeTo"), False) & "' "
If Request("CardNameFrom") <> "" Then sql = sql & " and T1.CardName >= N'" & saveHTMLDecode(Request("CardNameFrom"), False) & "' "
If Request("CardNameTo") <> "" Then sql = sql & " and T1.CardName <= N'" & saveHTMLDecode(Request("CardNameTo"), False) & "' "


If Request("GroupNameFrom") <> "" or Request("GroupNameTo") <> "" Then
	sql = sql & " and (( "
	
	If Request("GroupNameFrom") <> "" Then sql = sql & " T2.GroupName >= N'" & saveHTMLDecode(Request("GroupNameFrom"), False) & "' "
	If Request("GroupNameFrom") <> "" and Request("GroupNameTo") <> "" Then sql = sql & " and "
	If Request("GroupNameTo") <> "" Then sql = sql & " T2.GroupName <= N'" & saveHTMLDecode(Request("GroupNameTo"), False) & "' "
	
	sql = sql & ") or T2.GroupCode in (select PK " & _
					"	from OMLT X0 " & _
					"	inner join MLT1 X1 on X1.TranEntry = X0.TranEntry " & _
					"	where TableName = 'OCRG' and FieldAlias = 'GroupName' and LangCode = OLKCommon.dbo.OLKGetSBOLang(" & Session("LanID") & ") "
					
	If Request("GroupNameFrom") <> "" Then sql = sql & " and Convert(nvarchar(100),Trans) >= N'" & saveHTMLDecode(Request("GroupNameFrom"), False) & "' "
	If Request("GroupNameTo") <> "" Then sql = sql & " and Convert(nvarchar(100),Trans) <= N'" & saveHTMLDecode(Request("GroupNameTo"), False) & "' "
	
	sql = sql & ") ) "
End If

If Request("CountryFrom") <> "" Then sql = sql & " and T3.Name >= N'" & saveHTMLDecode(Request("CountryFrom"), False) & "' "
If Request("CountryTo") <> "" Then sql = sql & " and T3.Name <= N'" & saveHTMLDecode(Request("CountryTo"), False) & "' "

If Request("CardType") <> "" Then
	sql = sql & " and IsNull(T1.CardType, 'C') = '" & Request("CardType") & "' "
Else
	CardType = ""
	If myAut.HasAuthorization(45) Then CardType = "'C'"
	If myAut.HasAuthorization(78) Then 
		If CardType <> "" Then CardType = CardType & ", "
		CardType = CardType & "'S'"
	End If
	If myAut.HasAuthorization(77) Then 
		If CardType <> "" Then CardType = CardType & ", "
		CardType = CardType & "'L'"
	End If
	sql = sql & " and IsNull(T1.CardType, '') in (" & CardType & ") "
End If

If Request("orden1") <> "" Then 
	orden1 = Request("orden1")
	orden2 = Request("orden2")
Else
	orden1 = "LogNum"
	orden2 = "desc"
End If

sql = sql & "order by "

Select Case orden1
	Case "LogNum"
		sql = sql & "T0.LogNum "
	Case "CardCode"
		sql = sql & "T1.CardCode "
	Case "CardName"
		sql = sql & "T1.CardName "
	Case "DocDateSort"
		sql = sql & "T0.SubDate "
	Case Else
		sql = sql & Request("orden1") & " "
End Select

sql = sql & orden2


rs.open sql, conn, 3, 1
rs.PageSize = 10
rs.CacheSize = 10

If Request("p") <> "" Then iCurPage = CInt(Request("p")) Else iCurPage = 1
iPageCount = rs.PageCount
%>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
    <tr>
      <td>
      <img src="images/spacer.gif" width="100%" height="1" border="0" alt></td>
    </tr>
    <tr>
      <td bgcolor="#9BC4FF">
      <table border="0" cellpadding="0" bordercolor="#111111" width="100%" id="AutoNumber1">
        <tr>
          <td width="100%">
          <table style="width: 100%">
			<tr>
				<td><input type="button" name="btnSearch" value="<%=getcrdpenLngStr("DtxtSearch")%>" onclick="javascript:document.frmGo.cmd.value='searchClientPend';document.frmGo.submit();"></td>
				<td align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><%=getcrdpenLngStr("LtxtPendList")%></font></b></td>
			</tr>
			</table>
          </td>
        </tr>
		<% If Not rs.Eof Then
		rs.AbsolutePage = iCurPage
		LogNum = ""
		For i = 1 to rs.PageSize
			If i > 1 Then LogNum = LogNum & ", "
			LogNum = LogNum & rs("LogNum")
			rs.movenext
			If rs.eof then exit for
		Next
		rs.close
		
		sql = 	"select T0.LogNum, Convert(int,T0.SubDate), SubDate DocDate, CardCode, CardName, T0.Status " & _
		"from r3_obscommon..tlog T0 " & _
		"inner join r3_obscommon..TCRD T1 on T1.LogNum = T0.LogNum " & _
		"where T0.LogNum in (" & LogNum & ") " & _
		"order by "

		Select Case orden1
			Case "LogNum"
				sql = sql & "T0.LogNum "
			Case "CardCode"
				sql = sql & "T1.CardCode "
			Case "CardName"
				sql = sql & "T1.CardName "
			Case "DocDateSort"
				sql = sql & "2 "
			Case Else
				sql = sql & Request("orden1") & " "
		End Select

		sql = sql & orden2
		
		set rs = conn.execute(sql)
		
		do while not rs.eof %>
        <tr>
          <td width="100%">
          <table border="0" cellpadding="0" cellspacing="1"  bordercolor="#111111" width="100%" id="AutoNumber2">
            <tr>
              <td width="4%" bgcolor="#66A4FF">
    <a href="javascript:doGoCrd('<%=rs("LogNum")%>', '<%=rs("status")%>');">
    <img border="0" src="images/<%=Session("rtl")%>flechaselec.gif" align="left"></a></td>
              <td width="25%" bgcolor="#66A4FF"><font face="verdana" color="#000000" size="1"><%=RS("LogNum")%></font></td>
              <td width="71%" bgcolor="#66A4FF"><font face="verdana" color="#000000" size="1"><%=FormatDate(RS("DocDate"), True)%></font></td>
            </tr>
            <tr>
              <td width="4%">
		    <a href="javascript:delLogNum(<%=RS("LogNum")%>);">
		    <img border="0" src="images/remove.gif"></a></td>
              <td width="25%"><font size="1" face="verdana" color="#000000"><%=RS("CardCode")%></font></td>
              <td width="71%"><font size="1" face="verdana" color="#000000"><%=RS("CardName")%></font></td>
            </tr>
            </table>
          </td>
        </tr>
     <% 
     rs.movenext
	 loop %>
        <tr>
          <td width="100%" colspan="4">
			<table border="0" cellpadding="0" cellspacing="1" width="100%" id="table1" dir="ltr">
				<tr>
					<td width="16">
					<% If iCurPage > 1 Then %><a href="javascript:goP(<%=iCurPage-1%>);"><img border="0" src="images/flecha_prev.gif" width="16" height="16"></a><% Else %>&nbsp;<% End If %></td>
					<td>
					<p align="center">
					<select name="cmbPage" size="1" onchange="javascript:javascript:goP(this.value);">
					<% For i = 1 to iPageCount %>
					<option value="<%=i%>" <% If i = iCurPage Then %>selected<% End If %>><%=i%></option>
					<% Next %>
					</select></td>
					<td width="16">
					<% If iCurPage < iPageCount Then %><a href="javascript:goP(<%=iCurPage+1%>);"><img border="0" src="images/flecha_next.gif" width="16" height="16"></a><% End If %></td>
				</tr>
			</table>
			</td>
        </tr>
        <% Else %>
        <tr>
          <td width="100%" align="center"><b>
			<font size="1" face="verdana" color="#000000"><%=getcrdpenLngStr("DtxtNoData")%></font></b></td>
        </tr>
        <% End If %>
        <tr>
          <td width="100%"><hr color="#3385FF" size="1"></td>
        </tr>
        
        </table>
      </td>
    </tr>
    </table>
  </center>
</div>
<form name="frmGo" method="post" action="operaciones.asp">
<% For each itm in Request.Form
If itm <> "p" and itm <> "delLog" Then %>
<input type="hidden" name="<%=itm%>" value="<%=Request.Form(itm)%>">
<% End If 
Next %>
<% For each itm in Request.QueryString
If itm <> "p" Then %>
<input type="hidden" name="<%=itm%>" value="<%=Request.QueryString(itm)%>">
<% End If 
Next %>
<input type="hidden" name="p" value="<%=iCurPage%>">
<input type="hidden" name="delLog" value="">
</form>
<script language="javascript">
function goP(p)
{
	document.frmGo.cmd.value = 'pendClients';
	document.frmGo.p.value = p;
	document.frmGo.delLog.value = '';
	document.frmGo.submit();
}
function delLogNum(lognum)
{
	if (confirm('<%=getcrdpenLngStr("LtxtConfDelClient")%>'.replace('{0}', lognum)))
	{
		document.frmGo.cmd.value = 'pendClients';
		document.frmGo.delLog.value = lognum;
		document.frmGo.submit();
	}
}

function confReopen()
{
	return confirm('<%=getcrdpenLngStr("LtxtConfReOpen")%>')
}

function doGoCrd(logNum, Status)
{
	if (Status == 'H') if (!confReopen()) return;
	
	document.frmGoCrd.LogNum.value = logNum;
	document.frmGoCrd.status.value = Status;
	
	document.frmGoCrd.submit();
}
</script>
<form name="frmGoCrd" action="goCrd.asp" method="post">
<input type="hidden" name="LogNum" value="">
<input type="hidden" name="status" value="">
</form>
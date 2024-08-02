<!--#include file="lang/openActivities.asp" -->
<%
set rs = server.createobject("ADODB.RecordSet")

If Request("delLog") <> "" Then
	sql = "update R3_ObsCommon..TLOG set Status = 'B' where LogNum = " & Request("delLog")
	conn.execute(sql)
End If

If Request("rdDate2") = "" Then	rdDate2 = 2 Else rdDate2 = CInt(Request("rdDate2"))

If Not IsNull(myApp.AgentClientsFilter) Then
	AgentClientsFilter = " and T1.CardCode collate database_default not in (" & Replace(Replace(myApp.AgentClientsFilter, "@SlpCode", Session("vendid")), "@Type", 1) & ") "
	AgentClientsFilter = AgentClientsFilter & " and T1.CardCode collate database_default not in (" & Replace(Replace(myApp.AgentClientsFilter, "@SlpCode", Session("vendid")), "@Type", 5) & ") "
End If

ObjCode = ""

If myApp.EnableOCLG Then ObjCode = "33"

If ObjCode = "" Then ObjCode = "-1"

sourceType = Request("cmbSourceType")

sql = ""

sqlAddStr = ""

If Request("orden1") = "T1.CardCode" or Request("orden1") = "CardName" Then
	sqlAddStr = ", " & Request("orden1") & " collate database_default CardCode "
End If

If sourceType = "" or sourceType = "O" Then
	If Request("orden1") = "Action" Then sqlAddStr = ", T1.Action collate database_default Action "
	
	sql = sql & "select T0.LogNum TransNum, Convert(int,T1.Recontact), 'O' SourceType" & sqlAddStr & " " & _
				"from r3_obscommon..tlog T0 " & _
				"inner join r3_obscommon..TCLG T1 on T1.LogNum = T0.LogNum " & _
				"inner join OCRD T2 on T2.CardCode = T1.CardCode collate database_default " & _
				"left outer join ocry C2 on C2.code = T2.Country " & _
				"inner join ocrg C3 on C3.GroupCode = T2.GroupCode "
				
	If Request("SlpCodeFrom") <> "" or Request("SlpCodeTo") <> "" Then
		sql = sql & "inner join oslp S0 on S0.SlpCode = T1.SlpCode  " & _  
					"left outer join OMLT S1 on S1.TableName = 'OSLP' and S1.FieldAlias = 'SlpName' and S1.PK = S0.SlpCode " & _  
					"left outer join MLT1 S2 on S2.TranEntry = S1.TranEntry and S2.LangCode = OLKCommon.dbo.OLKGetSBOLang(" & Session("LanID") & ") " 
	End If
				
	sql = sql & "where T0.Object in (" & ObjCode & ") and T0.Status = 'R' and Company = '" & Session("olkDB") & "' and T1.ClgCode is null " & GetOpenActivitiesFilter("O") & " "
End If

If sourceType = "" Then sql = sql & "union "
		
If sourceType = "" or sourceType = "S" Then
	If Request("orden1") = "Action" Then sqlAddStr = ", Case T1.Action When 'N' Then 'O' Else T1.Action End collate database_default Action "
	
	sql = sql & "select T1.ClgCode TransNum, Convert(int,T1.Recontact), 'S' SourceType" & sqlAddStr & " " & _
				"from OCLG T1 " & _
				"inner join OCRD T2 on T2.CardCode = T1.CardCode collate database_default "& _
				"left outer join ocry C2 on C2.code = T2.Country " & _
				"inner join ocrg C3 on C3.GroupCode = T2.GroupCode " 
				
	If Request("SlpCodeFrom") <> "" or Request("SlpCodeTo") <> "" Then
		sql = sql & "left outer join OUSR T6 on T6.INTERNAL_K = T1.AttendUser " & _
					"left outer join OHEM T7 on T7.userId = T1.AttendUser " & _
					"left outer join oslp S0 on S0.SlpCode = T7.salesPrson  " & _  
					"left outer join OMLT S1 on S1.TableName = 'OSLP' and S1.FieldAlias = 'SlpName' and S1.PK = S0.SlpCode " & _  
					"left outer join MLT1 S2 on S2.TranEntry = S1.TranEntry and S2.LangCode = OLKCommon.dbo.OLKGetSBOLang(" & Session("LanID") & ") " 
	End If				
				
	sql = sql & "where 33 in (" & ObjCode & ") and Closed = 'N' and T1.Inactive = 'N' " & GetOpenActivitiesFilter("S") & " "
End If
		
sql = sql & "order by "

If Request("orden1") <> "" Then 
	orden1 = Request("orden1")
	orden2 = Request("orden2")
Else
	If myApp.ActOrdr1 <> "CntctDateSort" Then	orden1 = myApp.ActOrdr1 Else orden1 = "2"
	orden2 = myApp.ActOrdr2
End If

sql = sql & orden1 & " " & orden2

rs.open sql, conn, 3, 1
rs.PageSize = 10
rs.CacheSize = 10

sqlAddStr = ""

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
      <table border="0" cellpadding="0" bordercolor="#111111" width="100%">
        <tr>
          <td width="100%">
          <table style="width: 100%">
			<tr>
				<td><input type="button" name="btnSearch" value="<%=getopenActivitiesLngStr("DtxtSearch")%>" onclick="javascript:document.frmGo.cmd.value='openActivitiesSearch';document.frmGo.submit();"></td>
				<td align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><%=getopenActivitiesLngStr("LtxtOpenActivities")%></font></b></td>
			</tr>
			</table>
          </td>
        </tr>
        <tr>
          <td width="100%">
          <table style="width: 100%">
			<tr>
				<td style="width: 15px; "><input type="radio" name="rdDate2" id="rdDate2_0" value="0" <% If rdDate2 = 0 Then %>checked <% End If %> onclick="changeDate2(this.value);"></td>
				<td><label for="rdDate2_0"><b><font face="Verdana" size="1"><%=getopenActivitiesLngStr("DtxtToday")%></font></b></label></td>
				<td style="width: 15px; "><input type="radio" name="rdDate2" id="rdDate2_1" value="1" <% If rdDate2 = 1 Then %>checked <% End If %> onclick="changeDate2(this.value);"></td>
				<td><label for="rdDate2_1"><b><font face="Verdana" size="1"><%=getopenActivitiesLngStr("DtxtThisWeek")%></font></b></label></td>
				<td style="width: 15px; "><input type="radio" name="rdDate2" id="rdDate2_2" value="2" <% If rdDate2 = 2 Then %>checked <% End If %> onclick="changeDate2(this.value);"></td>
				<td><label for="rdDate2_2"><b><font face="Verdana" size="1"><%=getopenActivitiesLngStr("DtxtAll")%></font></b></label></td>
			</tr>
			</table>
          </td>
        </tr>
		<% If Not rs.Eof Then
		rs.AbsolutePage = iCurPage
		LogNum = ""
		ClgCode = ""
		For i = 1 to rs.PageSize
			Select Case rs("SourceType")
				Case "O"
					If LogNum <> "" Then LogNum = LogNum & ", "
					LogNum = LogNum & rs("TransNum")
				Case "S"
					If ClgCode <> "" Then ClgCode = ClgCode & ", "
					ClgCode = ClgCode & rs("TransNum")
			End Select
			rs.movenext
			If rs.eof then exit for
		Next
		rs.close
		
		sql = ""
		
		If (sourceType = "" or sourceType = "O") and LogNum <> "" Then
			sql = 	"select T0.LogNum TransNum, Convert(int,T1.Recontact), T1.Recontact, T1.CardCode collate database_default CardCode, " & _
					"OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRD', 'CardName', T1.CardCode, T2.CardName) collate database_default CardName, T1.Action collate database_default Action, " & _
					"T1.Details collate database_default Details, 'O' SourceType " & _
					"from r3_obscommon..tlog T0 " & _
					"inner join r3_obscommon..TCLG T1 on T1.LogNum = T0.LogNum " & _
					"inner join OCRD T2 on T2.CardCode = T1.CardCode collate database_default "& _
					"left outer join oslp T3 on T3.SlpCode = T1.SlpCode " & _
					"left outer join ocry C2 on C2.code = T2.Country " & _
					"inner join ocrg C3 on C3.GroupCode = T2.GroupCode " & _
					"left outer join OUSR T6 on T6.INTERNAL_K = T1.AttendUser " & _
					"where T0.LogNum in (" & LogNum & ") "
		End If
		
		If sourceType = "" and LogNum <> "" and ClgCode <> "" Then sql = sql & " union "
		
		If (sourceType = "" or sourceType = "S") and ClgCode <> "" Then
			sql = 	sql & "select T1.ClgCode TransNum, Convert(int,T1.Recontact), T1.Recontact, T1.CardCode, " & _
					"OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRD', 'CardName', T1.CardCode, T2.CardName) CardName, Case T1.Action When 'N' Then 'O' Else T1.Action End Action, T1.Details, 'S' SourceType " & _
					"from OCLG T1 " & _
					"inner join OCRD T2 on T2.CardCode = T1.CardCode collate database_default "& _
					"left outer join oslp T3 on T3.SlpCode = T1.SlpCode " & _
					"left outer join ocry C2 on C2.code = T2.Country " & _
					"inner join ocrg C3 on C3.GroupCode = T2.GroupCode " & _
					"left outer join OUSR T6 on T6.INTERNAL_K = T1.AttendUser " & _
					"where T1.ClgCode in (" & ClgCode & ") "
		End If
		
		sql = sql & "order by " & orden1 & " " & orden2
		
		set rs = conn.execute(sql)
		
		do while not rs.eof %>
        <tr>
          <td width="100%">
          <table border="0" cellpadding="0" cellspacing="1"  bordercolor="#111111" width="100%">
            <tr>
              <td width="4%" bgcolor="#66A4FF">
    <a href="javascript:doGoAct('<%=rs("SourceType")%>', <%=rs("TransNum")%>, '<%=Replace(myHTMLEncode(rs("CardCode")), "'", "\'")%>');">
    <img border="0" src="images/<%=Session("rtl")%>flechaselec.gif" align="left"></a></td>
              <td width="28%" bgcolor="#66A4FF"><font face="verdana" color="#000000" size="1"><%=RS("TransNum")%></font></td>
              <td width="38%" bgcolor="#66A4FF"><font face="verdana" color="#000000" size="1"><%=FormatDate(RS("Recontact"), True)%></font></td>
              <td width="30%" bgcolor="#66A4FF"><font face="verdana" color="#000000" size="1"><% Select Case rs("Action")
					Case "C"
						Response.Write getopenActivitiesLngStr("DtxtConv")
					Case "M"
						Response.Write getopenActivitiesLngStr("DtxtMeeting")
					Case "E"
						Response.Write getopenActivitiesLngStr("DtxtNote")
					Case "O"
						Response.Write getopenActivitiesLngStr("DtxtOther")
					Case "T"
						Response.Write getopenActivitiesLngStr("DtxtTask")
					End Select %></font></td>
            </tr>
            <tr>
              <td width="4%">
			  <% If rs("SourceType") = "O" Then %><a href="javascript:delLogNum(<%=RS("TransNum")%>);"><img border="0" src="images/remove.gif"></a><% End If %></td>
              <td width="28%"><font size="1" face="verdana" color="#000000"><a href="operaciones.asp?cmd=datos&card=<%=CleanItem(myHTMLEncode(RS("CardCode")))%>"><font color="#000000"><%=RS("CardCode")%></font></a></td>
              <td width="38%" colspan="2" style="width: 68%">
              <table cellpadding="0" cellspacing="0" border="0" style="width: 100%;">
              	<tr>
              		<td>
              		<font size="1" face="verdana" color="#000000"><%=RS("CardName")%>&nbsp;</font></td>
              		<td width="13"><img src="images/icon_activity_<%=rs("SourceType")%>.gif"></td>
              	</tr>
              </table></td>
            </tr>
            <% If Not IsNull(rs("Details")) Then %>
            <tr>
              <td width="100%" colspan="4"><font size="1" face="verdana" color="#000000">
				<%=rs("Details")%></font>
			</td>
			<% End If %>
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
					<select name="cmbPage" size="1" onchange="javascript:goP(this.value);">
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
			<font size="1" face="verdana" color="#000000"><%=getopenActivitiesLngStr("DtxtNoData")%></font></b></td>
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
If itm <> "p" and itm <> "delLog" and itm <> "rdDate2" Then %>
<input type="hidden" name="<%=itm%>" value="<%=Request.Form(itm)%>">
<% End If 
Next %>
<% For each itm in Request.QueryString
If itm <> "p" and itm <> "rdDate2" Then %>
<input type="hidden" name="<%=itm%>" value="<%=Request.QueryString(itm)%>">
<% End If 
Next %>
<input type="hidden" name="p" value="<%=iCurPage%>">
<input type="hidden" name="rdDate2" value="<%=rdDate2%>">
<input type="hidden" name="delLog" value="">
</form>
<script language="javascript">
function goP(p)
{
	document.frmGo.cmd.value = 'openActivities';
	document.frmGo.p.value = p;
	document.frmGo.delLog.value = '';
	document.frmGo.submit();
}
function changeDate2(value)
{
	document.frmGo.cmd.value = 'openActivities';
	document.frmGo.delLog.value = '';
	document.frmGo.rdDate2.value = value;
	document.frmGo.submit();
}

function delLogNum(lognum)
{
	if (confirm('<%=getopenActivitiesLngStr("LtxtConfDelAct")%>'.replace('{0}', lognum)))
	{
		document.frmGo.cmd.value = 'openActivities';
		document.frmGo.delLog.value = lognum;
		document.frmGo.submit();
	}
}

function doGoAct(sourceType, transNum, CardCode)
{
	switch (sourceType)
	{
		case 'O':
			document.doGoAct.LogNum.value = transNum;
			document.doGoAct.CardCode.value = CardCode;
			document.doGoAct.submit();
			break;
		case 'S':
			document.doGoEditAct.ClgCode.value = transNum;
			document.doGoEditAct.CardCode.value = CardCode;
			document.doGoEditAct.submit();
			break;
	}
}
</script>
<form name="doGoAct" action="goAct.asp" method="post">
<input type="hidden" name="LogNum" value="">
<input type="hidden" name="CardCode" value="">
</form>
<form name="doGoEditAct" action="goActEdit.asp" method="post">
<input type="hidden" name="ClgCode" value="">
<input type="hidden" name="CardCode" value="">
</form>
<%

Function GetOpenActivitiesFilter(ByVal FilterType)
	strFilter = ""
	
	If Session("useraccess") = "U" and not myAut.HasAuthorization(60) Then
		Select Case FilterType
			Case "O"
				strFilter = " and T1.SlpCode = " & Session("VendId")
			Case "S"
				strFilter = " and T7.salesPrson = " & Session("VendId")
		End Select
	End If	
	
	strFilter = strFilter & " " & AgentClientsFilter
	
	Select Case FilterType
		Case "O"
			fldTrans = "T0.LogNum"
			fldAction = "T1.Action"
		Case "S"
			fldTrans = "T1.ClgCode"
			fldAction = "Case T1.Action When 'N' Then 'O' Else T1.Action End "
	End Select

	If Session("useraccess") = "U" Then
		CardType = ""
		If myAut.HasAuthorization(23) Then CardType = "'C'"
		If myAut.HasAuthorization(75) Then
			If CardType <> "" Then CardType = CardType & ", "
			CardType = CardType & "'L'"
		End If
		strFilter = strFilter & " and T2.CardType in (" & CardType & ") "
	End If

	If Request("dtFrom") <> "" Then strFilter = strFilter & " and DateDiff(day,Convert(datetime,'" & SaveSqlDate(Request("dtFrom")) & "',120), T1.Recontact) >= 0 "
	If Request("dtTo") <> "" Then strFilter = strFilter & " and DateDiff(day,T1.Recontact, Convert(datetime,'" & SaveSqlDate(Request("dtTo")) & "',120)) >= 0"
	
	Select Case rdDate2 
		Case 0 
			strFilter = strFilter & " and DateDiff(day,T1.Recontact,getdate()) = 0 "
		Case 1
			strFilter = strFilter & " and DateDiff(week,T1.Recontact,getdate()) < 1 "
	End Select
	

	'If Request("SlpCodeFrom") <> "" Then strFilter = strFilter & " and T3.SlpName >= N'" & saveHTMLDecode(Request("SlpCodeFrom"), False) & "' "
	'If Request("SlpCodeTo") <> "" Then strFilter = strFilter & " and T3.SlpName <= N'" & saveHTMLDecode(Request("SlpCodeTo"), False) & "' "

	If Request("SlpCodeFrom") <> "" Then strFilter = strFilter & " and Convert(nvarchar(4000),IsNull(S2.Trans, S0.SlpName)) >= N'" & saveHTMLDecode(Request("SlpCodeFrom"), False) & "' "
	If Request("SlpCodeTo") <> "" Then strFilter = strFilter & " and Convert(nvarchar(4000),IsNull(S2.Trans, S0.SlpName)) <= N'" & saveHTMLDecode(Request("SlpCodeTo"), False) & "' "

	
	If Request("LogNumFrom") <> "" Then strFilter = strFilter & " and " & fldTrans & " >= " & Request("LogNumFrom") & " "
	If Request("LogNumTo") <> "" Then strFilter = strFilter & " and " & fldTrans & " <= " & Request("LogNumTo") & " "
	
	If Request("CardCodeFrom") <> "" Then strFilter = strFilter & " and T1.CardCode >= N'" & saveHTMLDecode(Request("CardCodeFrom"), False) & "' "
	If Request("CardCodeTo") <> "" Then strFilter = strFilter & " and T1.CardCode <= N'" & saveHTMLDecode(Request("CardCodeTo"), False) & "' "

	GroupCode = ""
	Country = ""
	
	If Request("GroupNameFrom") <> "" or Request("GroupNameTo") <> "" Then
		GroupCode = GroupCode & " and (( "
		
		If Request("GroupNameFrom") <> "" Then GroupCode = GroupCode & " C3.GroupName >= N'" & saveHTMLDecode(Request("GroupNameFrom"), False) & "' "
		If Request("GroupNameFrom") <> "" and Request("GroupNameTo") <> "" Then GroupCode = GroupCode & " and "
		If Request("GroupNameTo") <> "" Then GroupCode = GroupCode & " C3.GroupName <= N'" & saveHTMLDecode(Request("GroupNameTo"), False) & "' "
		
		GroupCode = GroupCode & ") or C3.GroupCode in (select PK " & _
						"	from OMLT X0 " & _
						"	inner join MLT1 X1 on X1.TranEntry = X0.TranEntry " & _
						"	where TableName = 'OCRG' and FieldAlias = 'GroupName' and LangCode = OLKCommon.dbo.OLKGetSBOLang(" & Session("LanID") & ") "
						
		If Request("GroupNameFrom") <> "" Then GroupCode = GroupCode & " and Convert(nvarchar(100),Trans) >= N'" & saveHTMLDecode(Request("GroupNameFrom"), False) & "' "
		If Request("GroupNameTo") <> "" Then GroupCode = GroupCode & " and Convert(nvarchar(100),Trans) <= N'" & saveHTMLDecode(Request("GroupNameTo"), False) & "' "
		
		GroupCode = GroupCode & ") ) "
	End If
	
	If Request("CountryFrom") <> "" or Request("CountryTo") <> "" Then
		Country = Country & " and (( "
		
		If Request("CountryFrom") <> "" Then Country = Country & " C2.Name >= N'" & saveHTMLDecode(Request("CountryFrom"), False) & "' "
		If Request("CountryFrom") <> "" and Request("CountryTo") <> "" Then Country = Country & " and "
		If Request("CountryTo") <> "" Then Country = Country & " C2.Name <= N'" & saveHTMLDecode(Request("CountryTo"), False) & "' "
		
		Country = Country & ") or C2.Code in (select PK " & _
						"	from OMLT X0 " & _
						"	inner join MLT1 X1 on X1.TranEntry = X0.TranEntry " & _
						"	where TableName = 'OCRY' and FieldAlias = 'Name' and LangCode = OLKCommon.dbo.OLKGetSBOLang(" & Session("LanID") & ") "
						
		If Request("CountryFrom") <> "" Then Country = Country & " and Convert(nvarchar(100),Trans) >= N'" & saveHTMLDecode(Request("CountryFrom"), False) & "' "
		If Request("CountryTo") <> "" Then Country = Country & " and Convert(nvarchar(100),Trans) <= N'" & saveHTMLDecode(Request("CountryTo"), False) & "' "
		
		Country = Country & ") ) "
	End If
	
	If Request("CardNameFrom") <> "" or Request("CardNameTo") <> "" Then
		CardName = CardName & " and (( "
		
		If Request("CardNameFrom") <> "" Then CardName = CardName & " T2.CardName >= N'" & saveHTMLDecode(Request("CardNameFrom"), False) & "' "
		If Request("CardNameFrom") <> "" and Request("CardNameTo") <> "" Then CardName = CardName & " and "
		If Request("CardNameTo") <> "" Then CardName = CardName & " T2.CardName <= N'" & saveHTMLDecode(Request("CardNameTo"), False) & "' "
		
		CardName = CardName & ") or T2.CardCode in (select PK " & _
						"	from OMLT X0 " & _
						"	inner join MLT1 X1 on X1.TranEntry = X0.TranEntry " & _
						"	where TableName = 'OCRD' and FieldAlias = 'CardName' and LangCode = OLKCommon.dbo.OLKGetSBOLang(" & Session("LanID") & ") "
						
		If Request("CardNameFrom") <> "" Then CardName = CardName & " and Convert(nvarchar(100),Trans) >= N'" & saveHTMLDecode(Request("CardNameFrom"), False) & "' "
		If Request("CardNameTo") <> "" Then CardName = CardName & " and Convert(nvarchar(100),Trans) <= N'" & saveHTMLDecode(Request("CardNameTo"), False) & "' "
		
		CardName = CardName & ") ) "
	End If
	
'	If Request("AttendUserFrom") <> "" Then strFilter = strFilter &" and OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OUSR', 'U_NAME', T6.INTERNAL_K, T6.U_Name) >= N'" & saveHTMLDecode(Request("AttendUserFrom"), False) & "' "
'	If Request("AttendUserTo") <> "" Then strFilter = strFilter & " and OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OUSR', 'U_NAME', T6.INTERNAL_K, T6.U_Name) <= N'" & saveHTMLDecode(Request("AttendUserTo"), False) & "' "
	
	If Request("DocType") <> "" Then 
		strFilter = strFilter & " and " & fldAction & " = '" & Request("DocType") & "'"
	End If
	
	strFilter = strFilter & GroupCode & Country & CardName
	
	GetOpenActivitiesFilter = strFilter
End Function
%>
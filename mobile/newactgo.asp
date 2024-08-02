<!--#include file="lang/newactgo.asp" -->
<% varx = 0 %>
<SCRIPT LANGUAGE="JavaScript">

<!-- Begin
function Start(page) {
OpenWin = this.open(page, "CtrlWindow", "toolbar=no,menubar=no,location=no,scrollbars=no,resizable=no, height=378,width=420");
}
// End -->
</SCRIPT>

<%

If Request.Form.Count > 0 Then
	OnlyMyAct = Request("chkMyAct") = "Y"
	ShowClosed = Request("ShowClosed") = "Y"
Else
	OnlyMyAct = True
	ShowClosed = False
End If

If Request("delLog") <> "" Then
	sql = "update R3_ObsCommon..TLOG set Status = 'B' where LogNum = " & Request("delLog")
	conn.execute(sql)
End If


sqlAdd = ""
sqlAdd2 = ""

If OnlyMyAct or not myAut.HasAuthorization(60) or not myAut.HasAuthorization(97) Then
	sqlAdd = " and T1.SlpCode = " & Session("vendid") & " "
	sqlAdd2 = " and T1.salesPrson = " & Session("vendid") & " "
End If

If Not ShowClosed Then sqlAdd2 = sqlAdd2 & " and T0.Closed = 'N' "

ObjCode = ""

If myApp.EnableOCLG Then ObjCode = "33"

If ObjCode = "" Then ObjCode = "-1"

sql = "declare @CardCode nvarchar(15) set @CardCode = N'" & saveHTMLDecode(Request("CardCode"), False) & "'   " & _  
"declare @OLKActCount int  " & _  
"set @OLKActCount = (select Count(T0.lognum)   " & _  
"				from r3_obscommon..tlog T0  " & _  
"				inner join r3_obscommon..TCLG T1 on T1.lognum = T0.lognum   " & _  
"				inner join OCRD T2 on T2.CardCode = T1.CardCode collate database_default  " & _  
"				where Company = N'" & Session("olkDB") & "' and T0.object in (" & ObjCode & ") " & sqlAdd & " and T0.status = 'R' and T1.cardcode = @CardCode) " & _  
"declare @SysActCount int " & _  
"set @SysActCount = (select Count(T0.ClgCode)   " & _  
"				from OCLG T0 " & _  
"				left outer join OHEM T1 on T1.userId = T0.AttendUser " & _
"				inner join OCRD T2 on T2.CardCode = T0.CardCode " & _  
"				where T0.cardcode = @CardCode and T0.Inactive = 'N' " & sqlAdd2 & ") " & _  
"select CardName, @OLKActCount+@SysActCount CountAct from OCRD where CardCode = @CardCode " 
set rs = conn.execute(sql)

sql = "select T0.LogNum TransNum, 'O' SourceType, T1.Recontact " & _
	  "from r3_obscommon..tlog T0 " & _
	  "inner join r3_obscommon..TCLG T1 on T1.lognum = T0.Lognum " & _
	  "where Company = '" & Session("olkDB") & "' and T0.object = 33 " & sqlAdd & " and T0.status = 'R' and T1.ClgCode is null and T1.cardcode = N'" & saveHTMLDecode(Request("CardCode"), False) & "' " & _
	  "union " & _
	  "select T0.ClgCode TransNum, 'S' SourceType, T0.Recontact " & _
	  "from OCLG T0 " & _
	  "left outer join OHEM T1 on T1.userId = T0.AttendUser " & _
	  "where T0.CardCode = N'" & saveHTMLDecode(Request("CardCode"), False) & "' and T0.Inactive = 'N' " & sqlAdd2 &  _
	  " order by 3 desc"
set rd = Server.CreateObject("ADODB.RecordSet")
rd.open sql, conn, 3, 1
rd.PageSize = 10
rd.CacheSize = 10

If Request("p") <> "" Then iCurPage = CInt(Request("p")) Else iCurPage = 1
iPageCount = rd.PageCount

If Not rd.Eof Then
	rd.AbsolutePage = iCurPage
	LogNum = ""
	ClgCode = ""
	For i = 1 to rd.PageSize
		Select Case rd("SourceType")
			Case "O"
				If LogNum <> "" Then LogNum = LogNum & ", "
				LogNum = LogNum & rd("TransNum")
			Case "S"
				If ClgCode <> "" Then ClgCode = ClgCode & ", "
				ClgCode = ClgCode & rd("TransNum")
		End Select
		rd.movenext
		If rd.Eof Then Exit For
	Next
	
	sql = ""
	If LogNum <> "" Then
		sql = "select T0.LogNum TransNum, 'O' SourceType, 'O' SourceTypeImg, T1.CardCode collate database_default CardCode, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRD', 'CardName', T1.CardCode, T2.CardName) CardName, " & _
			  "Recontact, Convert(int,T1.Recontact) RecontactSort, Details collate database_default Details, T1.Action collate database_default Action " & _
			  "from r3_obscommon..tlog T0 " & _
			  "inner join r3_obscommon..TCLG T1 on T1.lognum = T0.Lognum " & _
			  "inner join ocrd T2 on T2.cardcode = T1.cardcode collate database_default " & _
			  "where T0.LogNum in (" & LogNum & ") "
	End If
	
	If LogNum <> "" and ClgCode <> "" Then sql = sql & " union "
	
	If ClgCode <> "" Then
		sql = sql & "select T1.ClgCode TransNum, 'S' SourceType, Case T1.Closed When 'N' Then 'S' Else 'C' End SourceTypeImg, T1.CardCode, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRD', 'CardName', T1.CardCode, T2.CardName) CardName, " & _
				  "Recontact, Convert(int,T1.Recontact) RecontactSort, Details, Case T1.Action When 'N' Then 'O' Else T1.Action End Action " & _
				  "from OCLG T1  " & _
				  "inner join ocrd T2 on T2.cardcode = T1.cardcode collate database_default " & _
				  "where T1.ClgCode in (" & ClgCode & ") "
	End If

	sql = sql & "order by RecontactSort desc"
	
	set rd = conn.execute(sql)
End If

%><div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
    <tr>
      <td bgcolor="#9BC4FF">
      <table border="0" cellpadding="0" bordercolor="#111111" width="100%">
        <tr>
          <td width="100%">
          <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><%=getnewactgoLngStr("LtxtClientOpenAct")%> 
          </font></b></td>
        </tr>
        <tr>
          <td width="100%" bgcolor="#3385FF">
			<table border="0" width="100%" id="table1" cellspacing="0" cellpadding="0">
				<tr>
					<td width="12"><a href="operaciones.asp?cmd=datos&card=<%=CleanItem(Request("CardCode"))%>"><img border="0" src="images/<%=Session("rtl")%>flechaselec.gif" align="left"></a></td>
					<td><font face="Verdana" size="1" color="#000000">&nbsp;<%=Request("CardCode")%> - <%=rs("CardName")%></font></td>
				</tr>
			</table>
			</td>
        </tr>
        <tr>
          <td width="100%">
          <table border="0" cellpadding="0"  bordercolor="#111111" width="100%">
            <tr>
              <td width="100%"><b><font size="1" face="Verdana">
              <p align="center"><%=Replace(getnewactgoLngStr("LtxtPendingActs"), "{0}", rs("CountAct"))%></font></b></td>
            </tr>
            <% If myAut.HasAuthorization(67) Then %>
            <tr>
              <td width="100%">
              <p align="center">
              <input type="button" value="<%=getnewactgoLngStr("LtxtNewActivity")%>" onclick="javascript:window.location.href='newactgonow.asp?CardCode=<%=Replace(CleanItem(Request("CardCode")), "'", "\'")%>'">
              </td>
            </tr>
            <% End If %>
            <form name="frmMyAct" method="post" action="operaciones.asp">
	          <input type="hidden" name="cmd" value="<%=Request("cmd")%>">
	          <input type="hidden" name="CardCode" value="<%=Request("CardCode")%>">
            <% If myAut.HasAuthorization(97) Then %>
            <tr>
              <td width="100%"><font face="Verdana" size="1" color="#000000">
              <input type="checkbox" <% If OnlyMyAct Then %>checked<% End If %> name="chkMyAct" id="chkMyAct" value="Y" onclick="submit();">
              <label for="chkMyAct"><%=getnewactgoLngStr("LtxtShowMyAct")%></label></font>
              </td>
            </tr>
            <% End If %>
            <tr>
              <td width="100%"><font face="Verdana" size="1" color="#000000">
              <input type="checkbox" <% If ShowClosed Then %>checked<% End If %> name="ShowClosed" id="chkShowClosed" value="Y" onclick="submit();">
              <label for="chkShowClosed"><%=getnewactgoLngStr("LtxtShowClosedAct")%></label></font>
              </td>
            </tr>
            </form>
          </table>
          </td>
        </tr>
        <tr>
          <td width="100%"><hr color="#3385FF" size="1"></td>
        </tr>
        <% If Not rd.Eof Then
        do while not rd.eof %>
        <tr>
          <td width="100%">
          <table border="0" cellpadding="0" cellspacing="1"  bordercolor="#111111" width="100%">
            <tr>
              <td width="4%" bgcolor="#66A4FF">
    			<a href="javascript:doGoAct('<%=rd("SourceType")%>', <%=rd("TransNum")%>, '<%=Replace(myHTMLEncode(rd("CardCode")), "'", "\'")%>');">
    			<img border="0" src="images/<%=Session("rtl")%>flechaselec.gif" align="left"></a></td>
              <td width="50%" bgcolor="#66A4FF"><font face="verdana" color="#000000" size="1"><%=rd("TransNum")%></font></td>
              <td width="50%" bgcolor="#66A4FF"><font face="verdana" color="#000000" size="1"><%=FormatDate(rd("Recontact"), True)%></font></td>
            </tr>
            <tr>
              <td width="4%">
    			<% If rd("SourceType") = "O" Then %><a href="javascript:delLogNum(<%=rd("TransNum")%>);">
    			<img border="0" src="images/remove.gif"></a><% End If %></td>
              <td colspan="2">
              <table cellpadding="0" cellspacing="0" border="0" style="width: 100%;">
              	<tr>
              		<td>
              		<font size="1" face="verdana" color="#000000"><% Select Case rd("Action")
					Case "C"
						Response.Write getnewactgoLngStr("DtxtConv")
					Case "M"
						Response.Write getnewactgoLngStr("DtxtMeeting")
					Case "E"
						Response.Write getnewactgoLngStr("DtxtNote")
					Case "O"
						Response.Write getnewactgoLngStr("DtxtOther")
					Case "T"
						Response.Write getnewactgoLngStr("DtxtTask")
					End Select %>&nbsp;</font></td>
              		<td width="13"><img src="images/icon_activity_<%=rd("SourceTypeImg")%>.gif"></td>
              	</tr>
              </table></td>
            </tr>
            <% If Not IsNull(rd("Details")) Then %>
            <tr>
              <td width="100%" colspan="3"><font size="1" face="verdana" color="#000000">
				<%=rd("Details")%></font>
			</td>
			<% End If %>
            </tr>
            </table>
          </td>
        </tr>
	    <% rd.movenext
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
        <% End If %>
        <tr>
          <td width="100%">&nbsp;</td>
        </tr>
      </table>
      </td>
    </tr>
    </table>
  </center>
</div>
<form name="frmGo" method="post" action="operaciones.asp">
<% For each itm in Request.Form
If itm <> "p" and itm <> "delLog" and itm <> "chkMyAct" and itm <> "ShowClosed" Then %>
<input type="hidden" name="<%=itm%>" value="<%=Request.Form(itm)%>">
<% End If 
Next %>
<% For each itm in Request.QueryString
If itm <> "p" and itm <> "chkMyAct" and itm <> "ShowClosed" Then %>
<input type="hidden" name="<%=itm%>" value="<%=Request.QueryString(itm)%>">
<% End If 
Next %>
<input type="hidden" name="p" value="<%=iCurPage%>">
<input type="hidden" name="chkMyAct" value="<% If OnlyMyAct Then %>Y<% End If %>">
<input type="hidden" name="ShowClosed" value="<% If ShowClosed Then %>Y<% End If %>">
<input type="hidden" name="delLog" value="">
</form>
<script language="javascript">
function goP(p)
{
	document.frmGo.cmd.value = 'goActivities';
	document.frmGo.p.value = p;
	document.frmGo.delLog.value = '';
	document.frmGo.submit();
}
function doGoAct(sourceType, transNum, CardCode)
{
	switch (sourceType)
	{
		case 'O':
			document.frmGoAct.LogNum.value = transNum;
			document.frmGoAct.CardCode.value = CardCode;
			document.frmGoAct.submit();
			break;
		case 'S':
			document.doGoEditAct.ClgCode.value = transNum;
			document.doGoEditAct.CardCode.value = CardCode;
			document.doGoEditAct.submit();
			break;
	}
}

function delLogNum(lognum)
{
	if (confirm('<%=getnewactgoLngStr("LtxtConfDelAct")%>'.replace('{0}', lognum)))
	{
		document.frmGo.cmd.value = 'goActivities';
		document.frmGo.delLog.value = lognum;
		document.frmGo.submit();
	}
}
</script>
<form name="frmGoAct" action="goAct.asp" method="post">
<input type="hidden" name="LogNum" value="">
<input type="hidden" name="CardCode" value="">
</form>
<form name="doGoEditAct" action="goActEdit.asp" method="post">
<input type="hidden" name="ClgCode" value="">
<input type="hidden" name="CardCode" value="">
</form>
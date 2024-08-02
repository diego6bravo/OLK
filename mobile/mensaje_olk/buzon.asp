<% addLngPathStr = "mensaje_olk/" %>
<!--#include file="lang/buzon.asp" -->
<%

sql = "Select T0.OlkLog, Case OlkUFromType When 'V' Then (select OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OSLP', 'SlpName', SlpCode, SlpName) collate database_default from oslp where slpcode = OLKUFrom) When 'C' Then " & _
	  "(select OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRD', 'CardName', CardCode, CardName) from ocrd where cardcode = OLKUFrom) collate database_default When 'S' Then 'Sistema' End OlkUFrom, " & _
	  "Convert(nvarchar(30),OlkSubject) + Case When Len(OlkSubject) > 30 Then N'...' Else '' End as OlkSubject, OlkMSG, OlkUrgent, OlkStatus, " & _
	  "OlkDate Date " & _
	  "from olkomsg T0 " & _
	  "inner join olkmsg1 T1 on T1.olklog = T0.olklog " & _
	  "where OlkUser = '" & Session("vendid") & "' and OlkStatus <> 'D' " & _
	  "order by OlkDate desc"
rs.open sql, conn, 3, 1
rs.PageSize = 6
rs.CacheSize = 6

iPageCount = RS.PageCount

If Request("p") <> "" Then iCurPage = CInt(Request("p")) Else iCurPage = 1
If iCurPage > iPageCount Then iCurPage = iPageCount
If iCurPage < 1 Then iCurPage = 1
%>
<style>
a            { color: #0053CE }
</style>
<script type="text/javascript">
function valFrm()
{
	if (document.frmDel.delLog)
	{
		var found = false;
		if (document.frmDel.delLog.length)
		{
			for (var i = 0;i<document.frmDel.delLog.length;i++)
			{
				if (document.frmDel.delLog[i].checked)
				{
					found = true;
					break;
				}
			}
		}
		else
		{
			found = document.frmDel.delLog.checked;
		}
		
		if (found)
		{
			return confirm('<%=getbuzonLngStr("LtxtConfCancel")%>');
		}
		else
		{
			alert('<%=getbuzonLngStr("LtxtSelDelMsg")%>');
			return false;
		}
	}
	else
	{
		return false;
	}
}
</script>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
    <tr>
      <td bgcolor="#9BC4FF">
      <form method="post" action="mensaje_olk/delMessage.asp" name="frmDel" onsubmit="javascript:return valFrm();">
      <table border="0" cellpadding="0" bordercolor="#111111" width="100%" id="AutoNumber1">
        <tr>
          <td width="100%">
          <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><%=getbuzonLngStr("LtxtInbox")%></font></b></td>
        </tr>
        <tr>
          <td width="100%">
          <table border="0" cellpadding="0" cellspacing="1"  bordercolor="#111111" width="100%" id="AutoNumber2">
            <tr>
              <td align="center" colspan="3" bgcolor="#66A4FF"><b>
              <font size="1" face="Verdana"><%=getbuzonLngStr("DtxtDate")%></font></b></td>
              <td align="center" bgcolor="#66A4FF" colspan="2"><b><font size="1" face="Verdana">
              <%=getbuzonLngStr("DtxtUser")%></font></b></td>
            </tr>
            <tr>
              <td align="center" colspan="5" bgcolor="#66A4FF"><b><font size="1" face="Verdana">
              <%=getbuzonLngStr("DtxtSubject")%></font></b></td>
            </tr>
            <% 
            varx = 0
            If Not rs.Eof Then
            RS.AbsolutePage = iCurPage 
            for intRecord=1 to rs.PageSize
            varx = varx + 1
            If varx = 1 Then
            	bgColor = "#84B5FF"
            ElseIf varx = 2 Then
            	varx = 0
            	bgColor = "#AECEFF"
            End If %>
            <tr>
              <td bgcolor="<%=bgColor%>" width="14" <% If RS("olkUrgent") <> "Y" Then %>colspan="2"<% End If %>> <% If RS("OlkStatus") = "N" Then %><a href="mensaje_olk/updatestatus.asp?olklog=<%=RS("olklog")%>&status=O"><img border="0" src="images/mail_icon_new.gif"></a>
              <% ElseIf RS("olkstatus") = "O" Then %><a href="mensaje_olk/updatestatus.asp?olklog=<%=RS("olklog")%>&status=N"><img border="0" src="images/mail_icon_open.gif"></a><% end if %></td>
              <% If RS("olkUrgent") = "Y" Then %><td bgcolor="<%=bgColor%>"><img src="images/mail_icon_urgent.gif"></td><% end if %>
              <td bgcolor="<%=bgColor%>"><font size="1" face="Verdana"><a href="operaciones.asp?cmd=messageDetail&olklog=<%=RS("olklog")%>"><font color="#000000"><%=FormatDate(RS("Date"), True)%></font></a></td>
              <td bgcolor="<%=bgColor%>"><font size="1" face="Verdana"><a href="operaciones.asp?cmd=messageDetail&olklog=<%=RS("olklog")%>"><font color="#000000"><%=RS("OLKUFrom")%></font></a><font color="#000000">&nbsp;</font></td>
              <td bgcolor="<%=bgColor%>" width="16"><input type="checkbox" name="delLog" id="delLog" value="<%=RS("olklog")%>"></td>
            </tr>
            <tr>
              <td width="100%" bgcolor="<%=bgColor%>" colspan="5"><font size="1" face="Verdana"><a href="operaciones.asp?cmd=messageDetail&olklog=<%=RS("olklog")%>"><font color="#000000"><%=RS("OlkSubject")%></a>&nbsp;</font></td>
            </tr>
            <% 
            rs.movenext
			if rs.EOF then exit for
			next %>
            <tr>
              <td width="100%" colspan="5" align="center">
              <input type="submit" name="btnDel" value="<%=getbuzonLngStr("LtxtDelete")%>"></td>
            </tr>
			<% Else %>
            <tr>
              <td width="100%" colspan="5" align="center"><b>
				<font face="Verdana" size="1"><%=getbuzonLngStr("LtxtNoMsg")%></font></b></td>
            </tr>
            <% End If %>
            <tr>
              <td width="100%" colspan="5">
				<table border="0" cellpadding="0" cellspacing="1" width="100%" id="table1" dir="ltr">
					<tr>
						<td width="16">
						<% If iCurPage > 1 Then %><a href="operaciones.asp?cmd=buzon&amp;p=<%=iCurPage-1%>"><img border="0" src="images/flecha_prev.gif" width="16" height="16"></a><% Else %>&nbsp;<% End If %></td>
						<td>
						<p align="center">
						<select name="cmbPage" size="1" onchange="javascript:window.location.href='operaciones.asp?cmd=buzon&p=' + this.value">
						<% For i = 1 to iPageCount %>
						<option value="<%=i%>" <% If i = iCurPage Then %>selected<% End If %>><%=i%></option>
						<% Next %>
						</select></td>
						<td width="16">
						<% If iCurPage < iPageCount Then %><a href="operaciones.asp?cmd=buzon&amp;p=<%=iCurPage+1%>"><img border="0" src="images/flecha_next.gif" width="16" height="16"></a><% End If %></td>
					</tr>
				</table>
				</td>
            </tr>
            <tr>
              <td width="100%" colspan="5" align="center">
				<table border="0" cellpadding="0" cellspacing="1" width="100%" id="table2">
					<tr>
						<td width="50%" align="center"><a href="operaciones.asp?cmd=mssolk">
						<b><font face="Verdana" size="1"><%=getbuzonLngStr("LtxtOLKMsg")%></font></b></a></td>
						<td width="50%" align="center"><a href="operaciones.asp?cmd=msssbo">
						<b><font face="Verdana" size="1"><%=getbuzonLngStr("LtxtSBOMsg")%></font></b></a></td>
					</tr>
				</table>
				</td>
            </tr>
            </table>
          </td>
        </tr>
        </table>
      </form>
      </td>
    </tr>
    </table>
  </center>
</div>
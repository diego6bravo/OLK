<% addLngPathStr = "mensaje_olk/" %>
<!--#include file="lang/messagedetail.asp" -->
<%
sql = _
"declare @SlpCode int set @SlpCode = " & Session("vendid") & _
" Declare @OlkLog int set @OlkLog = " & Request("olklog") & _
" if (select OlkStatus From Olkmsg1 where olkuser = @SlpCode and olklog = @OlkLog and OlkUserType = '" & userType & "') = 'N' " & _
"Begin update olkmsg1 set olkstatus = 'O' where olkuser = @SlpCode and olklog = @OlkLog End "
conn.execute(sql)
sql = _
"Select T0.OlkLog, Case OlkUFromType When 'V' Then (select OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OSLP', 'SlpName', SlpCode, SlpName) from oslp where slpcode = OLKUFrom) When 'C' Then " & _
"(select OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRD', 'CardName', CardCode, CardName) from ocrd where cardcode = OLKUFrom) End OlkUFrom, OlkMSG, OlkSUbject, OlkUrgent, OlkStatus, " & _
"OlkDate from olkomsg T0 inner join olkmsg1 T1 on T1.OlkLog = T0.OlkLog " & _
"where olkuser = '" & Session("vendid") & "' and T0.olklog = " & Request("olklog")

set rs = conn.execute(sql)
%><div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
    <!-- fwtable fwsrc="Z:\topmanage\logos\originales\pocket_art.png" fwbase="pocket_artpieza1.gif" fwstyle="FrontPage" fwdocid = "742308039" fwnested=""0" -->
    <tr>
      <td bgcolor="#9BC4FF">
      <table border="0" cellpadding="0" bordercolor="#111111" width="100%" id="AutoNumber1">
        <tr>
          <td width="100%">
          <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><%=getmessagedetailLngStr("LttlMsgDetail")%></font></b></td>
        </tr>
        <tr>
          <td width="100%">
          <table border="0" cellpadding="0" cellspacing="1"  bordercolor="#111111" width="100%" id="AutoNumber2">
            <tr>
              <td width="33%" align="center" bgcolor="#66A4FF" colspan="3"><b>
              <font size="1" face="Verdana"><%=getmessagedetailLngStr("DtxtDate")%></font></b></td>
              <td width="33%" align="center" bgcolor="#66A4FF"><b>
              <font size="1" face="Verdana"><%=getmessagedetailLngStr("DtxtUser")%></font></b></td>
            </tr>
            <tr>
              <td width="4%">
              <p align="center"><b><font size="1" face="Verdana"></font></b>
              <% If RS("olkUrgent") = "Y" Then %><img border="0" src="images/mail_icon_urgent.gif"><% end if %></td>
              <td width="7%">
			<img border="0" src="images/mail_icon_<% If RS("OlkStatus") = "N" Then response.write "new" Else response.write "open" %>.gif"></td>
              <td width="22%">
              <b><font face="Verdana" size="1"><%=FormatDate(RS("OlkDate"), True)%></font></b></td>
              <td width="33%">
              <p align="center"><b><font size="1" face="Verdana"><%=RS("OLKUFrom")%></font></b></td>
            </tr>
            <tr>
              <td width="100%" colspan="4">
				<table border="0" cellpadding="0" cellspacing="1" width="100%" id="table1">
					<tr>
						<td width="55" valign="top">
							<table cellpadding="0" cellspacing="0" border="0">
								<tr>
									<td width="55" bgcolor="#66A4FF"><b><font face="Verdana" size="1"><%=getmessagedetailLngStr("DtxtSubject")%>:</font></b></td>
								</tr>
							</table>
						</td>
						<td><font size="1" face="Verdana"><%=RS("OlkSubject")%></font></td>
					</tr>
				</table>
				</td>
            </tr>
          </table>
          </td>
        </tr>
        <tr>
          <td width="100%">
          <p align="center">
         <textarea rows="10" name="S1" cols="25"><%=RS("OlkMsg")%></textarea></td>
        </tr>
      </table>
      </td>
    </tr>
    </table>
  </center>
</div>
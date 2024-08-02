<% addLngPathStr = "client/" %><!--#include file="lang/clientContacts.asp" -->
<%
sql = "select CardCode, IsNull(CardName, '') CardName, " & _
"(select Command from R3_ObsCommon..TLOG where LogNum = T0.LogNum) Command " & _
"from R3_ObsCommon..TCRD T0 " & _
"where T0.LogNum = " & Session("CrdRetVal")
set rs = conn.execute(sql)
isUpdate = rs("Command") = "U"
CardCode = rs("CardCode")
CardName = rs("CardName")
 %>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
  <form method="post" action="client/submitClient.asp" name="frmCrd">
    <tr>
      <td>
      <img src="images/spacer.gif" width="100%" height="1" border="0" alt></td>
    </tr>
    <tr>
      <td bgcolor="#9BC4FF">
      <table border="0" cellpadding="0" bordercolor="#111111" width="100%" id="AutoNumber1">
        <tr>
          <td width="100%">
          <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><% If Not isUpdate Then %><%=getclientContactsLngStr("LttlNewClient")%><% Else %><%=getclientContactsLngStr("LttlEditClient")%><% End If %> - <%=getclientContactsLngStr("LtxtContacts")%>
          </font></b></td>
        </tr>
        <tr>
          <td width="100%">
         <!--#include file="clientMenu.asp"--></td>
        </tr>
        <tr>
          <td width="100%">
            <table border="0" cellpadding="0"  bordercolor="#111111" width="100%" id="AutoNumber3">
            <tr>
              <td width="30%" bgcolor="#7DB1FF"><b>
              <font size="1" face="Verdana"><%=getclientContactsLngStr("DtxtCode")%></font></b></td>
              <td width="70%" bgcolor="#8CBAFF"><p>
              <font size="1" face="Verdana"><%=myHTMLEncode(CardCode)%></font></td>
            </tr>
			<tr>
              <td width="30%" bgcolor="#7DB1FF"><b>
              <font size="1" face="Verdana"><%=getclientContactsLngStr("DtxtName")%></font></b></td>
              <td width="70%" bgcolor="#8CBAFF"><p>
              <font size="1" face="Verdana"><%=myHTMLEncode(CardName)%></font></td>
            </tr>
          </table>
           </td>
        </tr>
        <tr>
          <td width="100%">
            <table border="0" cellpadding="0"  bordercolor="#111111" width="100%" id="AutoNumber3">
            <tr>
              <td bgcolor="#7DB1FF" width="15"></td>
              <td bgcolor="#7DB1FF"><b>
              <font size="1" face="Verdana"><%=getclientContactsLngStr("DtxtName")%></font></b></td>
              <td bgcolor="#7DB1FF"><b>
              <font size="1" face="Verdana"><%=getclientContactsLngStr("DtxtPosition")%></font></b></td>
            </tr>
            <%
				sql = "select T0.LineNum, T0.NewName, T0.Title, T0.Position, " & _
				"Case When T0.NewName = T1.CntctPrsn Then 'Y' Else 'N' End IsDefault " & _
				"from R3_ObsCommon..CRD2 T0 " & _
				"inner join R3_ObsCommon..TCRD T1 on T1.LogNum = T0.LogNum " & _
				"where T0.LogNum = " & Session("CrdRetVal") & " " & _
				"order by NewName"
				set rs = conn.execute(sql)
				If Not rs.Eof Then
				do while not rs.eof %>
			<tr>
              <td bgcolor="#8CBAFF" width="16"><a href='operaciones.asp?cmd=newClientContact&amp;EditID=<%=rs("LineNum")%>'><img border="0" src="images/<%=Session("rtl")%>flechaselec.gif" alt=""></a></td>
              <td bgcolor="#8CBAFF">
              <font size="1" face="Verdana"><% If rs("IsDefault") = "Y" Then %><b><% End If %><%=rs("NewName")%><% If rs("IsDefault") = "Y" Then %></b><% End If %></font></td>
              <td bgcolor="#8CBAFF"><p>
              <font size="1" face="Verdana"><% If rs("IsDefault") = "Y" Then %><b><% End If %><%=rs("Position")%><% If rs("IsDefault") = "Y" Then %></b><% End If %></font></td>
            </tr>
            <%	rs.movenext
            loop
            Else %>
			<tr>
              <td bgcolor="#8CBAFF" colspan="3" align="center"><p>
              <font size="1" face="Verdana"><%=getclientContactsLngStr("DtxtNoData")%></font></td>
            </tr>
            <% End If %>
			<tr>
              <td bgcolor="#8CBAFF" colspan="3"><input type="button" name="btnNew" value="<%=getclientContactsLngStr("DtxtNew")%>" onclick="window.location.href='?cmd=newClientContact';"></td>
            </tr>
          </table>
           </td>
        </tr>
        <tr>
          <td width="100%">&nbsp;</td>
        </tr>
      </table>
      </td>
    </tr>
	<input type="hidden" name="cmd" value="contacts">
	</form>
    </table>
  </center>
</div>
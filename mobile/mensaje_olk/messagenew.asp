<% addLngPathStr = "mensaje_olk/" %>
<!--#include file="lang/messagenew.asp" -->
<%
sql = "SELECT T0.SlpCode AS User_Code, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OSLP', 'SlpName', T0.SlpCode, T0.SlpName) U_Name " & _
	  "FROM OSLP T0 " & _
	  "inner join OLKAgentsAccess T1 on T1.SlpCode = T0.SlpCode and T1.Access <> 'D' " & _
	  "WHERE T0.SlpCode not in (-1," & Session("vendid") & ") " & _
	  "ORDER BY U_Name "
rs.open sql, conn, 3, 1
%><div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
    <tr>
      <td bgcolor="#9BC4FF">
      <table border="0" cellpadding="0" bordercolor="#111111" width="100%" id="AutoNumber1">
        <tr>
          <td width="100%">
          <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><%=getmessagenewLngStr("LtxtNewMessage")%>
          </font></b></td>
        </tr>
        <tr>
          <td width="100%">
          <script language="javascript">
          function valFrm()
          {
          	if (document.frmNewMsg.Subject.value == '')
          	{
          		alert('<%=getmessagenewLngStr("LtxtValSubject")%>');
          		document.frmNewMsg.Subject.focus();
          		return false;
          	}
          	if (!selUser())
          	{
          		alert('<%=getmessagenewLngStr("LtxtValSelUsr")%>');
          		return false;
          	}
          	else if (document.frmNewMsg.Message.value == '')
          	{
          		alert('<%=getmessagenewLngStr("LtxtValMsg")%>');
          		document.frmNewMsg.Message.focus();
          		return false;
          	}
          	return true;
          }
          function selUser()
          {
          	<% 
          	If Not rs.Eof Then 
          	do while not rs.eof %>
          	if (document.frmNewMsg.U<%=rs("USER_CODE")%>.checked) return true;
          	<% rs.movenext
          	loop
          	rs.movefirst
          	End If %>
          	return false;
          }
		  </script>
          <form method="POST" action="operaciones.asp" name="frmNewMsg" onsubmit="return valFrm();">
            <table border="0" cellpadding="0" cellspacing="1"  bordercolor="#111111" width="100%" id="AutoNumber2">
              <tr>
                <td width="100%" bgcolor="#66A4FF">
				<p align="center"><b><font size="1" face="Verdana"><%=getmessagenewLngStr("DtxtSubject")%></font></b></td>
              </tr>
              <tr>
                <td width="100%" bgcolor="#66A4FF">
				<p align="center"><b><font size="1" face="Verdana">
                <input name="Subject" size="32" maxlength="80"></font></b></td>
              </tr>
              <tr>
                <td width="100%" bgcolor="#66A4FF">
				<p align="center"><b><font size="1" face="Verdana"><%=getmessagenewLngStr("LtxtUsers")%></font></b></td>
              </tr>
              <tr>
                <td width="41%" valign="top">
                <table border="0" cellpadding="0" cellspacing="1"  bordercolor="#111111" width="100%" id="AutoNumber3">
                  <tr>
                  	<% do while not rs.eof
                  	If varx = 2 Then
                  		varx = 0
                  		Response.Write "</tr><tr>"
                  	End If %>
                    <td width="50%" bgcolor="#66A4FF">
					<input type="checkbox" name="U<%=RS("User_Code")%>" value="ON" id="fp<%=rs.bookmark%>"><font color="#000000" face="verdana" size="1"><label for="fp<%=rs.bookmark%>"><b><%=RS("U_Name")%></b></label></td>
					<% 
			        varx = varx + 1
			        rs.movenext
			        loop %>
                  </tr>
                </table>
                </td>
              </tr>
              <tr>
                <td width="100%" valign="top" bgcolor="#66A4FF">
                <p align="center"><b><font size="1" face="Verdana"><%=getmessagenewLngStr("DtxtMessage")%></font></b></td>
              </tr>
              <tr>
                <td width="100%" valign="top" bgcolor="#66A4FF">
                    <p align="center">
                    <textarea rows="9" name="Message" cols="26"></textarea></td>
              </tr>
              <tr>
                <td width="100%" valign="top" bgcolor="#66A4FF">
                <p align="center">
                    &nbsp;<table border="0" cellpadding="0" cellspacing="1" width="100%" id="table1">
					<tr>
						<td>
						<p align="center">
						<input type="checkbox" name="Urgent" value="Y" id="Urgent"><label for="Urgent"><font size="1" face="Verdana"><b><%=getmessagenewLngStr("LtxtUrgent")%></b></font></label></td>
						<td>
						<p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
                    <input type="submit" value="<%=getmessagenewLngStr("DtxtSend")%>" name="B1"> -
                    <input type="reset" value="<%=getmessagenewLngStr("DtxtClear")%>" name="B2">&nbsp; </td>
					</tr>
				</table>
				</td>
              </tr>
              </table>
          	<input type="hidden" name="cmd" value="olkMessagePost">
          </form>
          </td>
        </tr>
        </table>
      </td>
    </tr>
    </table>
  </center>
</div>
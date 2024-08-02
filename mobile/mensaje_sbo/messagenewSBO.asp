<% addLngPathStr = "mensaje_sbo/" %>
<!--#include file="lang/messagenewSBO.asp" -->
<%
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetSysUsers" & Session("ID")
cmd.Parameters.Refresh()
cmd("@LanID") = Session("LanID")
rs.open cmd, , 3, 1
%>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
    <tr>
      <td bgcolor="#9BC4FF">
      <table border="0" cellpadding="0" bordercolor="#111111" width="100%">
        <tr>
          <td width="100%">
          <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><%=getmessagenewSBOLngStr("LtxtNewMessage")%> 
          </font></b></td>
        </tr>
        <tr>
          <td width="100%">
          <script language="javascript">
          function valFrm()
          {
          	if (document.frmNewMsg.Subject.value == '')
          	{
          		alert('<%=getmessagenewSBOLngStr("LtxtValSubject")%>');

          		document.frmNewMsg.Subject.focus();
          		return false;
          	}
          	if (!selUser())
          	{
          		alert('<%=getmessagenewSBOLngStr("LtxtValSelUsr")%>');
          		return false;
          	}
          	else if (document.frmNewMsg.Message.value == '')
          	{
          		alert('<%=getmessagenewSBOLngStr("LtxtValMsg")%>');
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
          	if (document.frmNewMsg.fp<%=rs.bookmark%>.checked) return true;
          	<% rs.movenext
          	loop
          	rs.movefirst
          	End If %>
          	return false;
          }
          </script>
          <form method="POST" action="operaciones.asp" name="frmNewMsg" onsubmit="return valFrm();">
            <table border="0" cellpadding="0" cellspacing="1"  bordercolor="#111111" width="100%">
              <tr>
                <td width="100%" bgcolor="#66A4FF">
				<p align="center"><b><font size="1" face="Verdana"><%=getmessagenewSBOLngStr("DtxtSubject")%></font></b></td>
              </tr>
              <tr>
                <td width="100%" bgcolor="#66A4FF" align="center"><b><font size="1" face="Verdana">
                <input name="Subject" size="32"></font></b></td>
              </tr>
              <tr>
                <td width="100%" bgcolor="#66A4FF">
				<p align="center"><b><font size="1" face="Verdana"><%=getmessagenewSBOLngStr("LtxtUsers")%></font></b></td>
              </tr>
              <tr>
                <td width="41%" valign="top">
                <table border="0" cellpadding="0" cellspacing="1"  bordercolor="#111111" width="100%">
                  <tr>
                  	<% do while not rs.eof
                  	If varx = 2 Then
                  		varx = 0
                  		Response.Write "</tr><tr>"
                  	End If %>
                    <td width="50%" bgcolor="#66A4FF">
					<input type="checkbox" name="SapTo" value="<%=RS("Code")%>" id="fp<%=rs.bookmark%>"><font color="#000000" face="verdana" size="1"><label for="fp<%=rs.bookmark%>"><b><%=RS("Name")%></b></label></td>
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
                <p align="center"><b><font size="1" face="Verdana"><%=getmessagenewSBOLngStr("DtxtMessage")%></font></b></td>
              </tr>
              <tr>
                <td width="100%" valign="top" bgcolor="#66A4FF">
                    <textarea rows="9" name="Message" cols="26"></textarea></td>
              </tr>
              <tr>
                <td width="100%" valign="top" bgcolor="#66A4FF">
                <table border="0" cellpadding="0" cellspacing="1" width="100%" id="table1">
					<tr>
						<td>
						<input type="checkbox" name="Urgent" value="Y" id="Urgent"><label for="Urgent"><font size="1" face="Verdana"><b><%=getmessagenewSBOLngStr("LtxtUrgent")%></b></font></label></td>
						<td>
						<p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
                    <input type="submit" value="<%=getmessagenewSBOLngStr("DtxtSend")%>" name="btnSend"> -
                    <input type="reset" value="<%=getmessagenewSBOLngStr("DtxtClear")%>" name="btnClear">&nbsp; </td>
					</tr>
				</table>
				</td>
              </tr>
              </table>
          	<input type="hidden" name="cmd" value="sboMessagePost">
          </form>
          </td>
        </tr>
        </table>
      </td>
    </tr>
    </table>
  </center>
</div>
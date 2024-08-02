<% addLngPathStr = "inv/" %>
<!--#include file="lang/invact.asp" -->
<%
If Session("branch") <> -1 Then
	sql = "select WhsCode from OLKBranchs where BranchIndex = " & Session("branch")
	set rs = conn.execute(sql)
	Session("bodega") = rs(0)
	
	If Request("redir") = "inv" Then 
		redirCmd = "searchitem&btnSearch=Y"
	Else 
		redirCmd = Request("redir")
		If Request("redir") = "invChkInOut" Then redirCmd = redirCmd & "&Type=" & Request("Type")
	End If	
	Response.Redirect "operaciones.asp?cmd=" & redirCmd
End If


set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetWarehouses" & Session("ID")
cmd.Parameters.Refresh()
cmd("@LanID") = Session("LanID")
set rs = Server.CreateObject("ADODB.RecordSet")
rs.open cmd, , 3, 1
%><div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
    <tr>
      <td bgcolor="#9BC4FF">
      <table border="0" cellpadding="0" bordercolor="#111111" width="100%">
        <tr>
          <td width="100%">
          <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><% If Request("redir") = "inv" Then %><%=getinvactLngStr("LtxtInvRecount")%><% ElseIf Request("redir") = "invChkInOut" Then %><%=getinvactLngStr("LtxtDelSearchOrder")%><% End If %></font></b></td>
        </tr>
        <tr>
          <td width="100%">
          <p align="center"><font size="1" face="Verdana"><%=getinvactLngStr("LtxtSelWHS")%></font></td>
        </tr>
        <tr>
          <td width="100%">
          <form method="POST" name="frmSelWHS" action="operaciones.asp?cmd=<% If Request("redir") = "inv" Then %>searchitem<% Else %><%=Request("redir")%><% End If %><% If Request("redir") = "invChkInOut" Then %>&Type=<%=Request("Type")%><% End If %>" onsubmit="return valFrm();">
          <table border="0" cellpadding="0" cellspacing="1"  bordercolor="#111111" width="100%" id="AutoNumber2">
            <tr>
              <td width="100%">
              <p align="center">
        <select size="5" name="bodega">
        <% while not rs.eof %>
        <option value="<%=rs("WhsCode")%>"><%=rs("WhsCode")%>-<%=rs("WhsName")%></option>
        <% rs.movenext
        wend %>
        </select></td>
            </tr>
            <tr>
              <td width="100%">
              <p align="center">
              <% If Request("redir") = "inv" Then %>
		        <input type="submit" value="<%=getinvactLngStr("DbtnSearch")%>" name="btnSearch"> - 
				<input type="submit" value="<%=getinvactLngStr("LtxtRep")%>" name="btnRep">
				<% ElseIf Request("redir") = "invChkInOut" Then %>
				<input type="submit" value="<%=getinvactLngStr("LtxtSelect")%>" name="btnSelect">
				<% End If %></td>
            </tr>
          </table>
          </form>
          </td>
        </tr>
        <tr>
          <td width="100%">&nbsp;</td>
        </tr>
      </table>
      </td>
    </tr>
    </table>
  </center>
</div>
<script language="javascript">
function valFrm()
{
	<% If Request("redir") = "delivery" Then %>
	if (document.frmSelWHS.bodega.value == '')
	{
		alert('<%=getinvactLngStr("LtxtValSelHWS")%>');
		document.frmSelWHS.bodega.focus();
		return false;
	}
	<% End If %>
	return true;
}
</script>
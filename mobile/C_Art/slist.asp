<% addLngPathStr = "C_Art/" %>
<!--#include file="lang/slist.asp" -->
<%
response.buffer = true
Dim sap1
Dim sap2
Dim sap3
Dim sap4

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetPriceListFiltered" & Session("ID")
cmd.Parameters.Refresh()
cmd("@LanID") = Session("LanID")
cmd("@UserAccess") = Session("UserAccess")
cmd("@SlpCode") = Session("vendid")
set rd = cmd.execute()
%>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
    <tr>
      <td bgcolor="#9BC4FF">
      <table border="0" cellpadding="0" bordercolor="#111111" width="100%" id="AutoNumber1">
        <tr>
          <td width="100%">
          <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><%=getslistLngStr("DtxtCat")%>&nbsp;-&nbsp;<%=getslistLngStr("LtxtSelPList")%> 
          </font></b></td>
        </tr>
        <tr>
          <td width="100%">
          <% If Not rd.Eof Then %>
          <form method="POST" action="operaciones.asp?cmd=slistsearch" name="search1">
           
            <table border="0" cellpadding="0"  bordercolor="#111111" width="100%" id="AutoNumber2">
              <tr>
                <td width="100%" bgcolor="#66A4FF">
                <p align="center">
			    <select size="1" name="plist">
			    <% while not rd.eof %>
			    <option value="<%=RD("ListNum")%>"><%=myHTMLEncode(RD("ListName"))%></option>
			    <% rd.movenext
			    wend %>
			    </select></td>
              </tr>
              <tr>
                <td width="100%" bgcolor="#66A4FF">
                <p align="center">
                <input border="0" src="images/ok_icon.gif" name="I1" type="image"></td>
              </tr>
              </table>
          </form>
          <% Else %>
            <table border="0" cellpadding="0"  bordercolor="#111111" width="100%" id="AutoNumber2">
              <tr>
                <td width="100%" bgcolor="#66A4FF">
	          <font face="Verdana" size="1"><%=getslistLngStr("LtxtNoActivePList")%></font>
	         	</td>
	         	</tr>
	         	</table>
          <% End If %>
          </td>
        </tr>
        </table>
      </td>
    </tr>
    </table>
  </center>
</div>
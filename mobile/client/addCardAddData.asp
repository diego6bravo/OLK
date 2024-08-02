<% addLngPathStr = "client/" %><!--#include file="lang/addCardAddData.asp" -->
<%
sql = "select CardCode, IsNull(CardName, '') CardName, Phone1, Phone2, Cellular, Fax, E_Mail, Notes, " & _
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
          <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><% If Not isUpdate Then %><%=getaddCardAddDataLngStr("LttlNewClient")%><% Else %><%=getaddCardAddDataLngStr("LttlEditClient")%><% End If %> - <%=getaddCardAddDataLngStr("DtxtAddData")%>
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
              <font size="1" face="Verdana"><%=getaddCardAddDataLngStr("DtxtCode")%></font></b></td>
              <td width="70%" bgcolor="#8CBAFF"><p>
              <font size="1" face="Verdana"><%=myHTMLEncode(CardCode)%></font></td>
            </tr>
			<tr>
              <td width="30%" bgcolor="#7DB1FF"><b>
              <font size="1" face="Verdana"><%=getaddCardAddDataLngStr("DtxtName")%></font></b></td>
              <td width="70%" bgcolor="#8CBAFF"><p>
              <font size="1" face="Verdana"><%=myHTMLEncode(CardName)%></font></td>
            </tr>
			<tr>
              <td width="30%" bgcolor="#7DB1FF"><b>
              <font size="1" face="Verdana"><%=getaddCardAddDataLngStr("DtxtPhone")%> 1</font></b></td>
              <td width="70%" bgcolor="#8CBAFF">
					<input type="text" name="Phone1" size="35" value="<%=myHTMLEncode(rs("Phone1"))%>" maxlength="20"></td>
            	</tr>
				<tr>
              <td width="30%" bgcolor="#7DB1FF"><b>
              <font size="1" face="Verdana"><%=getaddCardAddDataLngStr("DtxtPhone")%> 2</font></b></td>
              <td width="70%" bgcolor="#8CBAFF">
              <p>
    		<input type="text" name="Phone2" size="35" value="<%=myHTMLEncode(rs("Phone2"))%>" maxlength="20"></td>
            	</tr>
				<tr>
              <td width="30%" bgcolor="#7DB1FF"><b>
              <font size="1" face="Verdana"><%=getaddCardAddDataLngStr("LtxtCelular")%></font></b></td>
              <td width="70%" bgcolor="#8CBAFF">
              <input type="text" name="Cellular" size="35" value="<%=myHTMLEncode(rs("Cellular"))%>" maxlength="20"></td>
            	</tr>
				<tr>
              <td width="30%" bgcolor="#7DB1FF"><b>
              <font size="1" face="Verdana"><%=getaddCardAddDataLngStr("DtxtFax")%></font></b></td>
              <td width="70%" bgcolor="#8CBAFF">
              <input type="text" name="Fax" size="35" value="<%=myHTMLEncode(rs("Fax"))%>" maxlength="20"></td>
            	</tr>
			<tr>
              <td width="30%" bgcolor="#7DB1FF"><b>
              <font size="1" face="Verdana"><%=getaddCardAddDataLngStr("DtxtEMail")%></font></b></td>
              <td width="70%" bgcolor="#8CBAFF">
              <input type="text" name="E_Mail" size="35" value="<%=myHTMLEncode(rs("E_Mail"))%>" maxlength="100"></td>
            	</tr>
				<tr>
              <td width="30%" bgcolor="#7DB1FF"><b>
              <font size="1" face="Verdana"><%=getaddCardAddDataLngStr("DtxtObservations")%></font></b></td>
              <td width="70%" bgcolor="#8CBAFF">
              <textarea rows="4" name="Notes" cols="28" maxlength="100"><%=myHTMLEncode(rs("Notes"))%></textarea></td>
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
	<input type="hidden" name="cmd" value="addData">
	</form>
    </table>
  </center>
</div>
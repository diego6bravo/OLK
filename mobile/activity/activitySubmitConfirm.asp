<% addLngPathStr = "activity/" %>
<!--#include file="lang/activitySubmitConfirm.asp" -->
<%
If Request("s") = "E" Then
	sql = "select ErrMessage from R3_ObsCommon..TLOG where LogNum = " & Session("ActRetVal")
ElseIf Request("s") = "S" Then
	sql = "select ObjectCode, (select ClgCode from R3_ObsCommon..TCLG where LogNum = T0.LogNum) ClgCode from R3_ObsCommon..TLOG T0 where LogNum = " & Session("ConfActRetVal")
End If
set rs = conn.execute(sql)

%><div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
    <tr>
      <td bgcolor="#9BC4FF">
      <table border="0" cellpadding="0" bordercolor="#111111" width="100%" id="AutoNumber1">
        <tr>
          <td width="100%" bgcolor="#75ACFF">
          <p align="center">
      <b>
      <font face="Verdana" size="1"><%=getactivitySubmitConfirmLngStr("LtxtActivityConf")%></font></b></p>
          </td>
        </tr>
        <tr>
          <td width="100%">&nbsp;</td>
         </tr>
           <% If Request("s") = "E" Then %> 
        <tr>
          <td width="100%" align="center">
			<img src="images/error_icon.gif">
		  </td>
        </tr>      
        <% End If %>
        <tr>
        <form name="frmOpenAct" action="goActEdit.asp" method="post">
          <td width="100%">
          <p align="center"><font size="1" face="Verdana"><% 
          Select Case Request("s")
          	Case "S" %>
          	<b><% If IsNull(rs("ClgCode")) Then 
          			Response.Write Replace(getactivitysubmitconfirmLngStr("LtxtActAddOK"), "{0}", rs("ObjectCode"))
          			ClgCode = rs("ObjectCode")
          		  Else
          			Response.Write Replace(getactivitysubmitconfirmLngStr("LtxtActUpdOK"), "{0}", rs("ClgCode"))
          			ClgCode = rs("ClgCode")
          		  End If %></b><br>
          	<input type="submit" name="btnOpenAct" value="<%=getactivitySubmitConfirmLngStr("LtxtOpenAct")%>">
          	<input type="hidden" name="ClgCode" value="<%=ClgCode%>">
          	<input type="hidden" name="CardCode" value="<%=Server.HTMLEncode(Session("username"))%>">
	     <% Case "E" %>
	        <b><%=getactivitysubmitconfirmLngStr("LtxtErrAddAct")%></b><br><%=rs("ErrMessage")%>
         <% End Select %></font></td>
          	</form>
        </tr>
        <% If Request("s") <> "E" and 1 = 2 Then %>
        <form name="frmViewDoc" method="post" action="../cart/cxcDocDetail.asp">
        <tr>
          <td width="100%" align="center">
			<input type="image" src="../cart/images/print_OLK.gif" border="0" align="middle" id="fp1"><b><font size="1" face="Verdana"><label for="fp1">|D:txtPrint|</label></font></b></td>
        </tr>
        <input type="hidden" name="DocType" value="33">
        <input type="hidden" name="DocEntry" value="<% If Request("s") = "S" Then %><%=rs("ObjectCode")%><% Else %><%=Session("ConfRetVal")%><% End IF %>">
        </form>
        <% ElseIf Request("s") = "E" Then %>
        <tr>
          <td width="100%" align="center">
			<input type="button" name="btnRetry" value="<%=getactivitySubmitConfirmLngStr("DtxtRetry")%>" onclick="javascript:window.location.href='operaciones.asp?cmd=activitySubmit&retry=Y'">
		  </td>
        </tr>     
        <% End If %>
      </table>
      </td>
    </tr>
    </table>
  </center>
</div>
<%
If Session("NotifyAdd") Then
	Session("NotifyAdd") = False
	sql = "EXEC OLKCommon..DBOLKObjAlert" & Session("ID") & " " & Session("ConfActRetVal") & ", " & Session("branch") & ", 'V', '" & getMyLng & "'"
	conn.execute(sql)
End If
	
'Session("RetVal") = ""
%>
<% addLngPathStr = "activity/" %>
<% If session("OLKDB") = "" Then response.redirect "../lock.asp" %>
<!--#include file="lang/activitySubmit.asp" -->
<%

If Request("retry") = "Y" Then
	sql = "update R3_ObsCommon..TLOG set Status = 'C', ErrCode = null, ErrMessage = null, ErrLng = '" & GetLangErrCode() & "' where LogNum = " & Session("ActRetVal")
	conn.execute(sql)
End If

sql = 	"select status, object, tlog.objectcode, confirm from r3_obscommon..tlog tlog " & _
		"inner join olkdocconf G0 on G0.objectcode = tlog.object " & _
		"where lognum = " & Session("ActRetVal")
set rs = conn.execute(sql)

If (rs("status") = "S" or rs("Status") = "E") or (rs("Status") = "C" or rs("Status") = "P") and Request("bg") = "Y" Then
	Session("NotifyAdd") = rs("status") = "S"
	If rs("Status") <> "E" Then
		Session("ConfActRetVal") = Session("ActRetVal")
		Session("ActRetVal") = ""
	End If
	
	If Request("bg") = "Y" Then
		sql = "update R3_ObsCommon..TLOGControl set Background = 'Y', tag = 'V', myLng = '" & getMyLng & "', SlpCode = " & Session("vendid") & ", ConfBranch = " & Session("branch") & " where LogNum = " & Session("ConfActRetVal")
		conn.execute(sql)
	End If
	
	If rs("Status") = "E" Then
		sql = "update R3_ObsCommon..TLOG set Status = 'R' where LogNum = " & Session("ActRetVal")
		conn.execute(sql)
	End If

	response.redirect "operaciones.asp?cmd=activitySubmitConfirm&s=" & rs("status")
ElseIf rs("status") = "R" then
	If Request("Confirm") = "Y" Then updStatus = "H" Else updStatus = "C"
	sqlclose = 	"update r3_obscommon..tlog set status = '" & updStatus & "', ErrLng = '" & GetLangErrCode() & "' where LogNum = " & Session("ActRetVal")
	conn.execute(sqlclose)
End If
%>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
    <tr>
      <td bgcolor="#9BC4FF">
      <table border="0" cellpadding="0" bordercolor="#111111" width="100%" id="AutoNumber1">
        <tr>
          <td width="100%">
          <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><%=getactivitySubmitLngStr("LtxtSubmitActivity")%></font></b></td>
        </tr>
        <tr>
          <td width="100%">
          <p align="center"><img border="0" src="cart/gear_rueda.gif"></td>
        </tr>
        <tr>
          <td width="100%">
          <p align="center"><b><font face="Verdana" size="1"><%=getactivitySubmitLngStr("LtxtPleaseWait")%></font></b></td>
        </tr>
		<tr class="GeneralTblBold2">
			<td>
			<p align="center"><input type="button" name="runInBg" value="<%=getactivitySubmitLngStr("DtxtRunInBG")%>" onclick="javascript:window.location.href='operaciones.asp?cmd=activitySubmit&bg=Y';"></p>
			</td>
		</tr>
        <tr>
          <td width="100%">
          <p align="center">&nbsp;</td>
        </tr>
      </table>
      </td>
    </tr>
    </table>
  </center>
</div>
<script language="JavaScript"><!--
setTimeout("top.location.href = 'operaciones.asp?cmd=activitySubmit'",3000);
//--></script>

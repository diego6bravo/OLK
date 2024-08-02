<% addLngPathStr = "client/" %>
<% If session("OLKDB") = "" Then response.redirect "../lock.asp" %>
<!--#include file="lang/clientsubmit.asp" -->
<%

sql = 	"select Status, Object, T0.ObjectCode, Confirm, T2.CardType " & _
		"from r3_obscommon..tlog T0 " & _
		"inner join OLKDocConf T1 on T1.ObjectCode = T0.Object " & _
		"inner join R3_ObsCommon..TCRD T2 on T2.LogNum = T0.LogNum " & _
		"where T0.LogNum = " & Session("CrdRetVal")
set rs = conn.execute(sql)

If myAut.GetCardProperty(rs("CardType"), "C") or Request("Confirm") = "Y" Then
	updStatus = "H"
Else
	updStatus = "C"
End If 


If (rs("status") = "S" or rs("status") = "H" or rs("Status") = "E") or (rs("Status") = "C" or rs("Status") = "P") and Request("bg") = "Y" Then

	If rs("status") = "S" Then Session("NotifyAdd") = True Else Session("NotifyAdd") = False
	If rs("Status") <> "E" Then
		Session("ConfRetVal") = Session("CrdRetVal")
		Session("CrdRetVal") = ""
	End If
	
	If Request("bg") = "Y" Then
		sql = "update R3_ObsCommon..TLOGControl set Background = 'Y', tag = 'V', myLng = '" & getMyLng & "', SlpCode = " & Session("vendid") & ", ConfBranch = " & Session("branch") & " where LogNum = " & Session("ConfRetVal")
		conn.execute(sql)
	End If
	
	If rs("Status") = "E" Then
		sql = "update R3_ObsCommon..TLOG set Status = 'R' where LogNum = " & Session("CrdRetVal")
		conn.execute(sql)
	End If

	response.redirect "operaciones.asp?cmd=newClientSubmitConfirm&s=" & rs("status")
ElseIf rs("status") = "R" then
	Select Case updStatus
		Case "H"
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKCreateUAFControl" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@UserType") = "V"
			cmd("@ExecAt") = "C1" 
			cmd("@ObjectEntry") = Session("CrdRetVal")
			cmd("@AgentID") = Session("vendid")
			cmd("@LanID") = Session("LanID")
			cmd("@branch") = Session("branch")
			cmd("@SetLogNumConf") = "Y"
			cmd.execute()
		Case "C"
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKStartLogProcess"
			cmd.Parameters.Refresh()
			cmd("@LogNum") = Session("CrdRetVal")
			cmd("@ErrCode") = GetLangErrCode()
			cmd.execute()
	End Select
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
          <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><%=getclientsubmitLngStr("LtxtSubmitBP")%></font></b></td>
        </tr>
        <tr>
          <td width="100%">
          <p align="center"><img border="0" src="cart/gear_rueda.gif"></td>
        </tr>
        <tr>
          <td width="100%">
          <p align="center"><b><font face="Verdana" size="1"><%=getclientsubmitLngStr("LtxtPleaseWait")%></font></b></td>
        </tr>
		<% If updStatus = "C" Then %>
		<tr class="GeneralTblBold2">
			<td>
			<p align="center"><input type="button" name="runInBg" value="<%=getclientsubmitLngStr("DtxtRunInBG")%>" onclick="javascript:window.location.href='operaciones.asp?cmd=newClientSubmit&bg=Y';"></p>
			</td>
		</tr>
		<% End If %>
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
setTimeout("top.location.href = 'operaciones.asp?cmd=newClientSubmit'",3000);
//--></script>

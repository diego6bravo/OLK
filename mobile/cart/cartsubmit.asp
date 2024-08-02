<% addLngPathStr = "cart/" %>
<% If session("OLKDB") = "" Then response.redirect "../lock.asp" %>
<!--#include file="lang/cartsubmit.asp" -->
<%

sql = 	"select status, object, tlog.objectcode, confirm, (select BreakDown from OLKDocConf where ObjectCode = tlog.[Object]) BreakDown " & _
		" from r3_obscommon..tlog tlog " & _
		"inner join olkdocconf G0 on G0.objectcode = tlog.object " & _
		"where lognum = " & Session("RetVal")
set rs = conn.execute(sql)
BreakDown = rs("BreakDown") = "Y"
ObjCode = rs("object")
If myAut.GetObjectProperty(rs("object"), "C") or Request("Confirm") = "Y" Then
	updStatus = "H"
Else
	updStatus = "C"
End If 

If (rs("status") = "S" or rs("status") = "H" or rs("Status") = "E") or (rs("Status") = "C" or rs("Status") = "P") and Request("bg") = "Y" Then
	If rs("object") = 17 Then
		object = txtOrdr
	ElseIf rs("object") = 23 Then
		object = txtOqut
	End If

	If rs("status") = "S" Then Session("NotifyAdd") = True Else Session("NotifyAdd") = False
	If rs("Status") <> "E" Then
		Session("ConfRetVal") = Session("RetVal")
		Session("RetVal") = ""
	End If
	
	If Request("bg") = "Y" Then
		sql = "update R3_ObsCommon..TLOGControl set Background = 'Y', tag = 'V', myLng = '" & getMyLng & "', SlpCode = " & Session("vendid") & ", ConfBranch = " & Session("branch") & " where LogNum = " & Session("ConfRetVal")
		conn.execute(sql)
	End If
	
	If rs("Status") = "E" Then
		sql = "update R3_ObsCommon..TLOG set Status = 'R' where LogNum = " & Session("RetVal")
		conn.execute(sql)
	End If

	response.redirect "operaciones.asp?cmd=cartSubmitConfirm&doc=" & object & "&s=" & rs("status")
ElseIf rs("status") = "R" then
	Series = myAut.GetObjectProperty(ObjCode, "S")
	Series2 = myAut.GetObjectProperty(48, "S2")
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKCartSetDocFinalData" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@LanID") = Session("LanID")
	cmd("@LogNum") = Session("RetVal")
	cmd("@UserType") = userType
	cmd("@branch") = Session("branch")
	cmd("@ObjectCode") = rs("object")
	If Series <> "NULL" Then cmd("@Series") = Series
	If Series2 <> "NULL" Then cmd("@Series2") = Series2
	cmd("@SumDec") = myApp.SumDec
	cmd("@SlpCode") = Session("vendid")
	cmd.execute()
	
	If updStatus = "H" Then
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKCreateUAFControl" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@UserType") = "V"
		cmd("@ExecAt") = "D3" 
		cmd("@ObjectEntry") = Session("RetVal")
		cmd("@AgentID") = Session("vendid")
		cmd("@LanID") = Session("LanID")
		cmd("@branch") = Session("branch")
		cmd("@SetLogNumConf") = "Y"
		cmd.execute()
	Else
		If rs("BreakDown") = "Y" Then
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKExecuteBreakDown" & Session("ID")
			cmd.Parameters.Refresh
			cmd("@LogNum") = Session("RetVal")
			cmd.execute()
			If cmd("@BreakDown") = "Y" Then
				set cmd = Server.CreateObject("ADODB.Command")
				cmd.ActiveConnection = connCommon
				cmd.CommandType = &H0004
				cmd.CommandText = "DBExecuteDocCons"
				cmd.Parameters.Refresh
				cmd("@LogNum") = Session("RetVal")
				cmd.execute()
				Session("NotifyAdd") = True
				Response.Redirect "operaciones.asp?cmd=cartBreakDown"
			End If
		End If

		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKStartLogProcess"
		cmd.Parameters.Refresh()
		cmd("@LogNum") = Session("RetVal")
		cmd("@ErrCode") = GetLangErrCode()
		cmd.execute()
	End If
End If
%>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
    <!-- fwtable fwsrc="Z:\topmanage\logos\originales\pocket_art.png" fwbase="pocket_artpieza1.gif" fwstyle="FrontPage" fwdocid = "742308039" fwnested=""0" -->
    <tr>
      <td bgcolor="#9BC4FF">
      <table border="0" cellpadding="0" bordercolor="#111111" width="100%" id="AutoNumber1">
        <tr>
          <td width="100%">
          <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><%=getcartsubmitLngStr("LtxtSubmitCart")%></font></b></td>
        </tr>
        <tr>
          <td width="100%">
          <p align="center"><img border="0" src="cart/gear_rueda.gif"></td>
        </tr>
        <tr>
          <td width="100%">
          <p align="center"><b><font face="Verdana" size="1"><%=getcartsubmitLngStr("LtxtPleaseWait")%></font></b></td>
        </tr>
		<% If updStatus = "C" Then %>
		<tr class="GeneralTblBold2">
			<td>
			<p align="center"><input type="button" name="runInBg" value="<%=getcartsubmitLngStr("DtxtRunInBG")%>" onclick="javascript:window.location.href='operaciones.asp?cmd=cartSubmit&bg=Y';"></p>
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
setTimeout("top.location.href = 'operaciones.asp?cmd=cartSubmit'",3000);
//--></script>

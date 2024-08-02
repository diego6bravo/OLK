<% addLngPathStr = "inv/" %>
<% If session("OLKDB") = "" Then response.redirect "../lock.asp" %>
<!--#include file="lang/delOrderCheckSubmit.asp" -->
<% 
RetVal = 0
doSubmit = False
showErr = False
showOk = False
Object = 0

If Request("submit") <> "Y" Then
	Session("NotifyAdd") = True
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandText = "DBOLKCheckIOSubmit" & Session("ID")
	cmd.CommandType = &H0004
	cmd.Parameters.Refresh
	cmd("@LogNum") = Session("IORetVal")
	cmd("@ObjectCode") = Session("ObjCode")
	cmd("@Type") = Session("Type")
	cmd("@DocNum") = Request("txtOrderNum")
	cmd("@ErrLng") = GetLangErrCode()
	cmd.execute
	RetVal = Session("IORetVal")
	doSubmit = True
Else
	RetVal = Request("RetVal")
	sql = 	"select Status, ErrCode, ErrMessage, Object, " & _
			"(select T1 from OLKDocConf where ObjectCode = (select ChkOp from OLKInOutSettings where ObjectCode = " & Session("ObjCode") & " and Type = '" & Session("Type") & "')) oTable, Draft " & _
			"from R3_ObsCommon..TLOG T0 " & _
			"where LogNum = " & RetVal
	set rs = conn.execute(sql)
	Status = rs("Status")
	Draft = rs("Draft") = "Y"
	If rs("Status") = "C" or rs("Status") = "P" Then
		doSubmit = True
	ElseIf rs("Status") = "E" Then
		showErr = True
	ElseIf rs("Status") = "S" Then
		showOk = True
		If Session("NotifyAdd") Then
			Session("NotifyAdd") = False
			sql = "EXEC OLKCommon..DBOLKObjAlert" & Session("ID") & " " & RetVal & ", " & Session("branch") & ", 'V', '" & myLng & "'"
			conn.execute(sql)
		End If
	End If
End If 
%>
<div align="center">
	<center>
	<table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
		<tr>
			<td bgcolor="#9BC4FF">
			<table border="0" cellpadding="0" bordercolor="#111111" width="100%">
		        <tr>
		          <td width="100%">
		          <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><!--#include file="delOrderTitle.asp"-->
		          </font></b></td>
		        </tr>
				<% If not showErr and not showOk Then %>
				<tr>
					<td width="100%">
					<p align='<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>'>
					<b><font face="Verdana" size="1"><%=getdelOrderCheckSubmitLngStr("LtxtUpdSBO")%>
					</font></b></p>
					</td>
				</tr>
				<tr>
					<td width="100%" bgcolor="#82B4FF">
					<p align='center'>
					<b><font face="Verdana" size="1"><%=getdelOrderCheckSubmitLngStr("LtxtPleaseWait")%>
					</font></b></p>
					</td>
				</tr>
				<tr>
					<td width="100%">
					<p align='center'><img border="0" src="cart/gear_rueda.gif"></td>
				</tr>
				<% ElseIf showErr Then %>
				<tr>
					<td width="100%">
					<p align='<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>'>
					<b><font face="Verdana" size="1"><%=getdelOrderCheckSubmitLngStr("LtxtErr")%>
					</font></b></p>
					</td>
				</tr>
				<tr>
					<td width="100%" bgcolor="#82B4FF">
					<p align='center'>
					<b><font face="Verdana" size="1"><%=rs("ErrCode")%> - <%=rs("ErrMessage")%>
					</font></b></p>
					</td>
				</tr>
				<tr>
					<td width="100%" bgcolor="#82B4FF" align="center">
					<input type="button" name="btnGoToCheck" value="<%=getdelOrderCheckSubmitLngStr("LtxtBackToCheck")%>" onclick="javascript:window.location.href='?cmd=invChkInOutCheck&txtOrderNum=<%=Request("txtOrderNum")%>'">
					</td>
				</tr>
				<% ElseIf showOk Then %>
				<tr>
					<td width="100%">
					<p align='<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>'>
					<b><font face="Verdana" size="1"><%=getdelOrderCheckSubmitLngStr("LtxtUpdSBO")%>
					</font></b></p>
					</td>
				</tr>
				<tr>
					<td width="100%" bgcolor="#82B4FF">
					<p align='center'>
					<b><font face="Verdana" size="1"><% 
					myObj = rs("Object")
					oTable = rs("oTable")
					If Draft Then oTable = "ODRF"
					sql = "select DocNum, Lower(OLKCommon.dbo.DBOLKGetAlterNameByObjCode" & Session("ID") & "(" & Session("LanID") & ", " & myObj & ", 'S')) ObjDesc from " & oTable & " " & _
					"where DocEntry = (select ObjectCode from R3_ObsCommon..TLOG where LogNum = " & Request("RetVal") & ")"
					set rs = conn.execute(sql)
					If myObj = 17 Then
						Response.Write Replace(getdelOrderCheckSubmitLngStr("LtxtUpdORDR"), "{0}", rs("DocNum"))
					Else
						objDesc = rs("ObjDesc")
						If Draft Then 
							objDesc = objDesc & " (" & getdelOrderCheckSubmitLngStr("DtxtDraft") & ")"
						End If
						Response.Write Replace(Replace(getdelOrderCheckSubmitLngStr("LtxtAddDoc"), "{0}", objDesc), "{1}", rs("DocNum"))
					End If %>
					</font></b></p>
					</td>
				</tr>	
				<tr>
					<td width="100%">
					<p align='center'>
					<% Select Case Session("Type")
						Case "I"
							strLink = "?cmd=inv&redir=invChkInOut&Type=I"
						Case "O" 
							strLink = "?cmd=inv&redir=invChkInOut&Type=O"
					End Select %>
					<input type="button" name="btnNew" value="<%=getdelOrderCheckSubmitLngStr("DtxtNew")%>" onclick="javascript:window.location.href='<%=strLink%>';">
					</td>
				</tr>	

				<% End If %>
				<tr>
					<td height="5px"></td>
				</tr>
			</table>
			</td>
		</tr>
	</table>
	</center>
</div>
<% If doSubmit Then %>
<script language="JavaScript">
<!--
setTimeout("top.location.href = 'operaciones.asp?cmd=<%=Request("cmd")%>&txtOrderNum=<%=Request("txtOrderNum")%>&RetVal=<%=RetVal%>&submit=Y'",3000);
//-->
</script>
<% End If %>
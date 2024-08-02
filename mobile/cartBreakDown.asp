<% addLngPathStr = "" %>
<!--#include file="lang/cartBreakDown.asp" -->
<% If Session("NotifyAdd") Then
	Session("NotifyAdd") = False
	If Request("Consolidate") = "Y" Then
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKCartBreakDownCons" & Session("ID")
		cmd.Parameters.Refresh
		cmd("@LogNum") = Session("RetVal")
		cmd.execute()
	End If
	Session("ConfRetVal") = Session("RetVal")
	Session("RetVal") = ""
End If

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKCartGetBreakDown" & Session("ID")
cmd.Parameters.Refresh
cmd("@LogNum") = Session("ConfRetVal")
cmd("@LanID") = Session("LanID")
set rs = Server.CreateObject("ADODB.RecordSet")
rs.open cmd, , 3, 1 %>
<table border="0" cellpadding="0" width="100%">
	<tr class="GeneralTlt">
		<td><%=getcartBreakDownLngStr("LtxtDocBrekDown")%></td>
	</tr>
</table>
<table border="0" cellpadding="0" width="100%">
<form name="frmCheck">
	<tr class="GeneralTblBold2" style="text-align: center;">
		<td><%=getcartBreakDownLngStr("DtxtLogNum")%></td>
		<td><%=getcartBreakDownLngStr("DtxtBP")%></td>
		<td><%=getcartBreakDownLngStr("DtxtName")%></td>
		<td><%=getcartBreakDownLngStr("DtxtTotal")%></td>
		<td><%=getcartBreakDownLngStr("DtxtStatus")%></td>
	</tr>
	<% do while not rs.eof
	If rs("Status") = "C" or rs("Status") = "P" Then doRefresh = True %>
	<tr class="GeneralTbl">
		<td align="right"><%=rs("LogNum")%></td>
		<td><%=rs("CardCode")%></td>
		<td><%=rs("CardName")%></td>
		<td style="text-align: right;"><%=rs("DocCur")%>&nbsp;<%=FormatNumber(rs("DocTotal"), myApp.SumDec)%></td>
		<td align="center">
		<% Select Case rs("Status")
			Case "R", "C", "P"
				strImg = "ajax-loader.gif"
			Case "S"
				strImg = "check.gif" 
			Case "E"
				strImg = "error_db.gif"
		End Select %><img height="16" src="images/<%=strImg%>" width="16" id="StatusImg<%=rs("LogNum")%>">
		<input type="hidden" id="Status<%=rs("LogNum")%>" value="<%=rs("Status")%>">
		<input type="hidden" name="LogNum" value="<%=rs("LogNum")%>">
		</td>
	</tr>
	<% rs.movenext
	loop %>
</form>
</table><% If doRefresh Then %>
<script language="JavaScript"><!--
setTimeout("top.location.href = 'operaciones.asp?cmd=cartBreakDown'",3000);
//--></script><% End If %>
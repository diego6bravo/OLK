<% addLngPathStr = "messages/" %>
<!--#include file="lang/messagePost.asp" -->
<%
msgPostscriptname=Request.ServerVariables("SCRIPT_NAME")
msgPostxmlfilename=mid(msgPostscriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "messagePost.xml"
set msgPostdocmessagePost = server.CreateObject("MSXML2.DOMDocument")
msgPostdocmessagePost.async = False
msgPostDocmessagePost.Load(server.MapPath(msgPostxmlfilename)) 
msgPostdocmessagePost.setProperty "SelectionLanguage", "XPath"

function msgPostgetmessagePostLngStr(instring, lng)
	set msgPostselectedmessagePostnode = msgPostdocmessagePost.selectSingleNode("/languages/language[@xml:lang='" & lng & "']") 
	set msgPostselectedmessagePostnodes=msgPostdocmessagePost.documentElement.selectNodes("/languages/language")
	msgPosttemp = msgPostselectedmessagePostnode.selectSingleNode(instring).text
	msgPostgetmessagePostLngStr = msgPosttemp
end function


set rs = Server.CreateObject("ADODB.recordset")

If Not Session("sendmsg") Then
	If Request("AgentsTo") <> "" or Request("ClientsTo") <> "" Then
		
		Select Case userType 
			Case "V" 
				user = Session("vendid")
			Case "C"
				user = Session("UserName")
		End Select
		
		If Request("Urgente") = "Y" Then Urgent = "Y" Else Urgent = "N"
		
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.CommandText = "DBOlkIntMsg" & Session("ID")
		cmd.CommandType = &H0004
		cmd.ActiveConnection = connCommon
		cmd.Parameters.Refresh()
		cmd("@OlkUFrom") = saveHTMLDecode(user, False)
		cmd("@OlkUFromType") = userType
		cmd("@OlkMSG") = saveHTMLDecode(Request("Message"), True)
		cmd("@OlkSubject") = saveHTMLDecode(Request("Subject"), True)
		cmd("@OlkUrgent") = Urgent
		cmd.execute()
		olkLog = cmd("@OlkLog")
		
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.CommandText = "DBOLKIntMSGU" & Session("ID")
		cmd.CommandType = &H0004
		cmd.ActiveConnection = connCommon
		cmd.Parameters.Refresh()
		cmd("@OlkLog") = olkLog
		cmd("@OlkStatus") = "N"
		cmd("@LanID") = Session("LanID")
		
		If Request("AgentsTo") <> "" Then
			cmd("@UserType") = "V"
			ArrVal = Split(saveHTMLDecode(Request("AgentsTo"), True),", ")
			For i = 0 to UBound(ArrVal)
				cmd("@User") = ArrVal(i)
				cmd.execute()
			next
		End If
		
		If Request("ClientsTo") <> "" Then
			cmd("@UserType") = "C"
			ArrVal = Split(saveHTMLDecode(Request("ClientsTo"), True),", ")
			For i = 0 to UBound(ArrVal)
				cmd("@User") = ArrVal(i)
				cmd.execute()
			next
		End If

	End If
	If Request("SapTo") <> "" Then
		ArrVal = Split(saveHTMLDecode(Request("SapToName"), False),", ")
		strUsers = ""
		For i = 0 to UBound(ArrVal)
			If strUsers <> "" Then strUsers = strUsers & ", "
			strUsers = strUsers & "N'" & ArrVal(i) & "'"
		Next
		sql = "select IsNull(T1.AlertLang, T2.NatLng) LanID, T0.USERID,  " & _
			"T0.USER_CODE " & _
			"from OUSR T0 " & _
			"left outer join OLKAlertsTo T1 on T1.AlertType = 'S' and T1.AlertID = 2 and T1.ToType = 'S' and T1.ToID = T0.USERID " & _
			"cross join OLKCommon T2 " & _
			"where OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OUSR', 'U_NAME', T0.USERID, T0.U_NAME) collate database_default in (" & strUsers & ") " & _
			"order by 1, 2 "
		set ru = Server.CreateObject("ADODB.RecordSet")
		ru.open sql, conn, 3, 1
		
		sql = "select IsNull(T1.AlertLang, T2.NatLng) LanID, " & _
			"OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "( " & _
			"(select LanID from OLKCommon..OLKLang where Lower(Left(LanSign collate database_default, 2)) = IsNull(T1.AlertLang, T2.NatLng)), 'OSLP', 'SlpName', T3.SlpCode, T3.SlpName) SlpName " & _
			"from OUSR T0 " & _
			"left outer join OLKAlertsTo T1 on T1.AlertType = 'S' and T1.AlertID = 2 and T1.ToType = 'S' and T1.ToID = T0.USERID " & _
			"cross join OLKCommon T2 " & _
			"inner join OSLP T3 on T3.SlpCode = " & Session("vendid") & " " & _
			"where OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OUSR', 'U_NAME', T0.USERID, T0.U_NAME) collate database_default in (" & strUsers & ") "
		set rs = Server.CreateObject("ADODB.RecordSet")
		rs.open sql, conn, 3, 1
	
		set conn3 = Server.CreateObject("ADODB.Connection")
		set cmd = Server.CreateObject("ADODB.Command")
		conn3.Provider = olkSqlProv
		conn3.Open  "Provider=SQLOLEDB;charset=utf8;" & _
		          "Data Source=" & olkip & ";" & _
		          "Initial Catalog=R3_ObsCommon;" & _
		          "Uid=" & olklogin & ";" & _
		          "Pwd=" & olkpass & ""
		set var1 = Server.CreateObject("ADODB.recordset")
	      
		db = Session("olkdb")
		Set cmd.ActiveConnection = conn3
		cmd.CommandType = &H0004
		cmd.CommandText = "OBSSp_Request"
		cmd.Parameters.Refresh
		
		do while not rs.eof
			cmd.Execute , Array(0, db, Null, 81, "A", Null)
			RetVal2 = cmd.Parameters.Item(0).Value
			OBSSp_Request = RetVal2
			If Request("Urgente") = "Y" Then varPriority = 2 Else varPriority = 1
			strMsg = saveHTMLDecode(Request.Form("Message"), False)
			strMsg = msgPostgetmessagePostLngStr("DtxtFrom", rs("LanID")) & ": " & rs("SlpName") & VbCrLf & _
					msgPostgetmessagePostLngStr("DtxtMessage", rs("LanID")) & ": " & strMsg
				sql = "insert into tmsg(LogNum, Priority, Subject, UserText) " & _
						"values('" & RetVal2 & "', " & varPriority & ", " & _
						"N'" & saveHTMLDecode(Request.Form("Subject"), False) & "', N'" & strMsg & "')"
			conn3.execute(sql)
		
			sql = ""
			ru.Filter = "LanID = '" & rs("LanID") & "'"
			do while not ru.eof
				sql = sql & "insert into msg1(LogNum, UserCode, SendIntrnl) " & _
				"values('" & RetVal2 & "', N'" & ru("USER_CODE") & "', 'Y') "
			ru.movenext
			loop
			'response.redirect "../../query.asp?query=" & sql
			conn3.execute(sql)
			sql = "update tlog set status = 'C', ErrLng = '" & GetLangErrCode() & "' where lognum = " & RetVal2
			conn3.execute(sql)
		rs.movenext
		loop
	
		conn3.close
	End If
	Session("sendmsg") = True
End If

%>
<div align="center">
	<table border="0" cellpadding="0" width="350">
		<tr>
			<td>
			<p align="center">
			<img border="0" src="design/<%=SelDes%>/images/send_msg_img.gif" width="300" height="112"></td>
		</tr>
		<tr class="MsgTlt2">
			<td>
			<p align="center"><%=getmessagePostLngStr("LtxtOkMsg")%></td>
		</tr>
		<tr class="Msgtbl">
			<td>
			<p align="center"><a class="LinkNoticiasMas" href="<% If userType = "V" Then %>newMessage.asp<% Else %>messageReturn.asp?cmd=newMsg<% End If %>">
			<%=getmessagePostLngStr("LtxtSendOtherMsg")%></a></td>
		</tr>
		</table>
</div>
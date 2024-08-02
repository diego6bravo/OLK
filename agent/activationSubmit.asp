<!--#include file="clientInc.asp"-->
<!--#include file="agentTop.asp"-->
<% If userType <> "V" Then Response.Redirect "default.asp" %>
<% 
If Not (EnableClientActivation and myAut.HasAuthorization(89)) Then Response.Redirect "unauthorized.asp"
addLngPathStr = "" %>
<!--#include file="lang/activationSubmit.asp" -->
<% If Request("submit") = "Y" Then
	Session("NotifyAdd") = True
	Session("ConfRetVal")
	set conn2=Server.CreateObject("ADODB.Connection")
	conn2.Provider=olkSqlProv
	conn2.Open  "Provider=SQLOLEDB;charset=utf8;" & _
	          "Data Source=" & olkip & ";" & _
	          "Initial Catalog=R3_ObsCommon;" & _
	          "Uid=" & olklogin & ";" & _
	          "Pwd=" & olkpass & ""
	set Cmd = Server.CreateObject("ADODB.Command")
	Set cmd.ActiveConnection = conn2
	cmd.CommandType = adCmdStoredProc
	cmd.CommandText = "OBSSp_Request"
	cmd.Parameters.Refresh
	db = Session("olkdb")

	sql = "select T0.DocEntry from OCRD T0 " & _
	"inner join OLKClientsAccess T3 on T3.CardCode = T0.CardCode " & _
	"where T3.Status = 'P' "
	set rs = conn.execute(sql)
	sql = "declare @CardCode nvarchar(15) declare @LogNum int " & _
	"declare @Query nvarchar(4000)set @Query = N'insert R3_ObsCommon..tcrd(LogNum, CardCode, SlpCode) " & _
	"select @LogNum, CardCode, @SlpCode from OCRD where DocEntry = @DocEntry' "
	
	ConfDocEntry = ""
	do while not rs.eof
		DocEntry = rs("DocEntry")
		Status = Request("Status" & DocEntry)
		If Status <> "P" Then
		
			If ConfDocEntry <> "" Then ConfDocEntry = ConfDocEntry & ", "
			ConfDocEntry = ConfDocEntry & DocEntry
			
			If Status = "A" Then
				cmd.Execute , Array(0, db, Null, 2, "U", Null)
				RetVal = cmd.Parameters.Item(0).Value
				sql = sql & "set @LogNum = " & RetVal & " "

				If Session("ConfRetVal") <> "" Then Session("ConfRetVal") = Session("ConfRetVal") & ", "
				Session("ConfRetVal") = Session("ConfRetVal") & RetVal
			End If
			
			If Request("note" & DocEntry) <> "" and Status = "R" Then note = "N'" & saveHTMLDecode(Request("note" & DocEntry), False) & "'" Else note = "NULL"
			
			If Status = "R" Then
				sql = sql & "update OLKClientsAccess set Status = '" & Status & "', StatusDate = getdate(), " & _
							"StatusUserSign = " & Session("vendid") & ", StatusNote = " & Note & " " & _
							"where CardCode = (select CardCode collate database_default from OCRD where DocEntry = " & DocEntry & ") "
			End If
			If Status = "A" Then
				sql = sql & "EXEC sp_executesql @Query, N'@LogNum int, @SlpCode smallint, @DocEntry int', @LogNum = @LogNum, @SlpCode = " & Request("Slp" & DocEntry) & ", @DocEntry = " & DocEntry & " " & _
				"update R3_ObsCommon..tlog set status = 'C', ErrLng = '" & GetLangErrCode() & "' where lognum = @LogNum "
			End If
		End If
	rs.movenext
	loop
	conn.execute(sql)
	Wait
Else
	If Session("ConfRetVal") = "" Then
		Response.Redirect "activationConfirm.asp?ConfDocEntry=" & Request("ConfDocEntry")
	Else
		ConfDocEntry = Request("ConfDocEntry")
		sql = "select Status from R3_ObsCommon..TLOG where LogNum in (" & Session("ConfRetVal") & ") and Status in ('C', 'P')"
		set rs = conn.execute(sql)
		If Not rs.Eof Then
			Wait
		Else
			Response.Redirect "activationConfirm.asp?ConfDocEntry=" & Request("ConfDocEntry")
		End If
	End If
End If %>
<% Sub Wait %>
	
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
	<META HTTP-EQUIV=Refresh content="3; URL=activationSubmit.asp?ConfDocEntry=<%=ConfDocEntry%>">
	<title>Cart Submit</title>
	</head>

	<body>
	<p align="center">
	&nbsp;</p>
	<p align="center">
	&nbsp;</p>
	<p align="center">
	&nbsp;</p>
	<p align="center">
	&nbsp;</p>
	<p align="center">
	<table border="0" cellpadding="0" width="250" align="center">
		<tr>
			<td>
			<p align="center">
			<img border="0" src="design/<%=SelDes%>/images/gear.gif"></td>
		</tr>
		<tr class="FirmTlt">
			<td>
			<p align="center"><%=getactivationSubmitLngStr("DtxtWait")%>...</td>
		</tr>
	</table>
	<p align="center">&nbsp;</p>
	<p align="center">&nbsp;</p>
	<p align="center">&nbsp;</p>
	</body>
<% End Sub %>
<!--#include file="agentBottom.asp"-->
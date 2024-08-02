<!--#include file="../myHTMLEncode.asp"-->
<% If session("OLKDB") = "" Then response.redirect "../lock.asp" %>
<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

If Not Session("ActReadOnly") Then
	set rsdf = Server.CreateObject("ADODB.RecordSet")
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKGetUDFSystemCols" & Session("ID")
	cmd.Parameters.Refresh
	cmd("@LanID") = Session("LanID")
	cmd("@UserType") = userType
	cmd("@TableID") = "OCLG"
	cmd("@OP") = "P"
	rsdf.open cmd, , 3, 1

	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandText = "DBOLKCheckObjectCheckSum" & Session("ID")
	cmd.CommandType = &H0004
	cmd.Parameters.Refresh()
	cmd("@LogNum") = Session("ActRetVal")
	cmd.Execute()
	
	If cmd("@IsValid").Value = "N" Then
		Session("RetryRetVal") = Session("ActRetVal")
		Session("ActRetVal") = ""
		Response.Redirect "../operaciones.asp?cmd=crcerror"
	End If
	
	Select Case Request("cmd")
		Case "main"
			If Request("CntctCode") <> "" Then CntctCode = CInt(Request("CntctCode")) Else CntctCode = "NULL"
			If Request("Tel") <> "" Then Tel = "N'" & saveHTMLDecode(Request("Tel"), False) & "'" Else Tel = "NULL"
			If Request("CntctSbjct") <> "" Then CntctSbjct = "N'" & Request("CntctSbjct") & "'" Else CntctSbjct = "NULL"
			If Request("AttendUser") <> "" Then AttendUser = Request("AttendUser") Else AttendUser = "NULL"
			If Request("Details") <> "" Then Details = "N'" & saveHTMLDecode(Request("Details"), False) & "'" Else Details = "NULL"
			If Request("personal") = "Y" Then personal = "'Y'" Else personal = "'N'"
			If Request("Closed") = "Y" Then Closed = "'Y'" Else Closed = "'N'"
			Action = "'" & Request("Action") & "'"
			CntctType = Request("CntctType")
			Priority = "'" & Request("Priority") & "'"
			
			If Request("changeContact") = "Y" Then Tel = "(select IsNull(Case When T0.Tel1 is null or RTrim(T0.Tel1) = '' Then T1.Phone1 Else T0.Tel1 End, '') from OCPR T0 inner join OCRD T1 on T1.CardCode = T0.CardCode where T0.CntctCode = " & CntctCode & ") "
			
			rsdf.Filter = "FieldID = -1"
			If rsdf.Eof Then CntctCode = "CntctCode"
			rsdf.Filter = "FieldID = -2"
			If rsdf.Eof Then Tel = "Tel"
			rsdf.Filter = "FieldID = -3"
			If rsdf.Eof Then Action = "Action"
			rsdf.Filter = "FieldID = -4"
			If rsdf.Eof Then CntctType = "CntctType"
			rsdf.Filter = "FieldID = -5"
			If rsdf.Eof Then CntctSbjct = "CntctSbjct"
			rsdf.Filter = "FieldID = -6"
			If rsdf.Eof Then AttendUser = "AttendUser"
			rsdf.Filter = "FieldID = -7"
			If rsdf.Eof Then Priority = "Priority"
			rsdf.Filter = "FieldID = -8"
			If rsdf.Eof Then Details = "Details"
			rsdf.Filter = "FieldID = -9"
			If rsdf.Eof Then personal = "personal"
			rsdf.Filter = "FieldID = -10"
			If rsdf.Eof Then Closed = "Closed"

			sql = "update R3_ObsCommon..TCLG set CntctCode = " & CntctCode & ", Tel = " & Tel & ", Action = " & Action & ", CntctType = " & CntctType & ", CntctSbjct = " & CntctSbjct & ", " & _
					"AttendUser = " & AttendUser & ", Priority = " & Priority & ", Details = " & Details & ", personal = " & personal & ", Closed = " & Closed & " " & _
					"where LogNum = " & Session("ActRetVal")
			conn.execute(sql)
			
		Case "notes"
			If Request("Notes") <> "" Then Notes = "N'" & saveHTMLDecode(Request("Notes"), False) & "'" Else Notes = "NULL"
			sql = "update R3_ObsCommon..TCLG set Notes = " & Notes & " where LogNum = " & Session("ActRetVal")
			conn.execute(sql)
		Case "udf"
			sql = "select AliasID, TypeID, (select SDKID collate database_default from r3_obscommon..tcif where companydb = '" & Session("OLKDb") & "')++AliasID As InsertID " & _
				  "from cufd T0 " & _
				  "left outer join OLKCUFD T1 on T1.TableID = T0.TableID and T1.FieldID = T0.FieldID " & _
				  "where T0.TableId = 'OCLG' and AType in ('" & userType & "','T') and OP in ('T','P') and Active = 'Y' and IsNull(T1.GroupID, -1) = " & Request("GroupID")
			set rs = server.createobject("ADODB.RecordSet")
			rs.open sql, conn, 3, 1
			If Not rs.eof then
				sql = "update r3_obscommon..TCLG set "
				do while not rs.eof
					If rs.bookmark <> 1 Then sql = sql & ", "
					strVal = saveHTMLDecode(Request("U_" & rs("AliasID")), False)
					If Request("U_" & rs("AliasID")) <> "" Then 
						If rs("TypeID") = "D" Then AliasVal = "Convert(datetime,'" & SaveSqlDate(strVal) & "',120)" Else AliasVal = "N'" & strVal & "'" 
					Else 
						AliasVal = "NULL"
					End If
					sql = sql & rs("InsertID") & " = " & AliasVal
				rs.movenext
				loop
				sql = sql & " where lognum = " & Session("ActRetVal")
				conn.execute(sql)
			End If
		Case "address"
			If Request("Country") <> "" Then Country = "N'" & Request("Country") & "'" Else Country = "NULL"
			If Request("State") <> "" Then State = "N'" & Request("State") & "'" Else State = "NULL"
			If Request("City") <> "" Then City = "N'" & Request("City") & "'" Else City = "NULL"
			If Request("Street") <> "" Then Street = "N'" & Request("Street") & "'" Else Street = "NULL"
			If Request("Room") <> "" Then Room = "N'" & Request("Room") & "'" Else Room = "NULL"
			sql = "update R3_ObsCommon..TCLG set Country = " & Country & ", State = " & State & ", City = " & City & ", Street = " & Street & ", Room = " & Room & " where LogNum = " & Session("ActRetVal")
			conn.execute(sql)
		Case "general"
			If Request("Recontact") <> "" Then Recontact = "Convert(datetime,'" & SaveSqlDate(Request("Recontact")) & "',120)" Else Recontact = "NULL"
			If Request("endDate") <> "" Then endDate = "Convert(datetime,'" & SaveSqlDate(Request("endDate")) & "',120)" Else endDate = "NULL"
			If Request("Duration") <> "" Then Duration = Request("Duration") Else Duration = "NULL"
			If Request("Reminder") = "Y" Then Reminder = "Y" Else Reminder = "N"
			If Request("RemQty") <> "" Then RemQty = Request("RemQty") Else RemQty = "NULL"
			If Request("tentative") = "Y" Then tentative = "Y" Else tentative = "N"
			If Request("Inactive") = "Y" Then Inactive = "Y" Else Inactive = "N"
			If Request("DocEntry") <> "" Then DocEntry = Request("DocEntry") Else DocEntry = "NULL"
			Status = "'" & Request("Status") & "'"
			Location = "'" & Request("Location") & "'"
			DurType = "'" & Request("DurType") & "'"
			BeginTime = "OLKCommon.dbo.OLKConvertTimeToDate(" & Request("BeginTimeH") & "," & Request("BeginTimeM") & ",'" & Request("BeginTimeS") & "')"
			ENDTime = "OLKCommon.dbo.OLKConvertTimeToDate(" & Request("ENDTimeH") & "," & Request("ENDTimeM") & ",'" & Request("ENDTimeS") & "')"
			Inactive = "'" & Inactive & "'"
			RemType = "'" & Request("RemType") & "'"
			tentative = "'" & tentative & "'"
			DocType = Request("DocType")
			
			rsdf.Filter = "FieldID = -14"
			If rsdf.Eof Then Recontact = "Recontact"
			If rsdf.Eof Then BeginTime = "BeginTime"
			rsdf.Filter = "FieldID = -16"
			If rsdf.Eof Then endDate = "endDate"
			If rsdf.Eof Then ENDTime = "ENDTime"
			rsdf.Filter = "FieldID = -17"
			If rsdf.Eof Then Duration = "Duration"
			If rsdf.Eof Then DurType = "DurType"
			rsdf.Filter = "FieldID = -19"
			If rsdf.Eof Then Reminder = "Reminder"
			If rsdf.Eof Then RemType = "RemType"
			If rsdf.Eof Then RemQty = "RemQty"
			rsdf.Filter = "FieldID = -22"
			If rsdf.Eof Then tentative = "tentative"
			rsdf.Filter = "FieldID = -23"
			If rsdf.Eof Then Inactive = "Inactive"
			rsdf.Filter = "FieldID = -25"
			If rsdf.Eof Then DocEntry = "DocEntry"
			If rsdf.Eof Then DocType = "DocType"
			rsdf.Filter = "FieldID = -11"
			If rsdf.Eof Then Status = "Status"
			rsdf.Filter = "FieldID = -12"
			If rsdf.Eof Then Location = "Location"
			
			sql = "update R3_ObsCommon..TCLG set Status = " & Status & ", Location = " & Location & ", Duration = " & Duration & ", " & _
					"DurType = " & DurType & ", Reminder = " & Reminder & ", RemQty = " & RemQty & ", RemType = " & RemType & ", tentative = " & tentative & ", " & _
					"Inactive = " & Inactive & ", DocType = " & DocType & ", DocEntry = " & DocEntry & ", Recontact = " & Recontact & ", " & _
					"BeginTime = " & BeginTime & ", " & _
					"endDate = " & endDate & ", ENDTime = " & ENDTime & " " & _
					"where LogNum = " & Session("ActRetVal")
			conn.execute(sql)
	End Select
	
	If Request("btnUpdate.x") <> "" or Request("btnAdd.x") <> "" Then
		AddErr = getActAddError()
		If AddErr <> "" Then
			retURL = ""
			For each itm in Request.Form
				If LCase(itm) <> "item" and itm <> "Confirm" Then
					If retURL <> "" Then retURL = retURL & "{a}"
					retURL = retURL & itm & "{e}" & Request(itm)
				End If
			Next
			For each itm in Request.QueryString
				If LCase(itm) <> "item" and itm <> "Confirm" Then
					If retURL <> "" Then retURL = retURL & "{a}"
					retURL = retURL & itm & "{e}" & Request(itm)
				End If
			Next
			response.redirect "../operaciones.asp?cmd=DocFlowErr&DocFlowErr=" & AddErr & "&retURL=" & retURL
		End If
	End If
	
	TransType = "U"
	If Request("btnAdd.x") <> "" Then TransType = "A"
	
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKPostObjectCreation" & Session("ID")
	cmd.Parameters.Refresh
	cmd("@sessiontype") = "A"
	cmd("@transtype") = TransType
	cmd("@object") = 33 
	cmd("@LogNum") = Session("ActRetVal")
	cmd("@CurrentSlpCode") = Session("vendid")
	cmd("@Branch") = Session("branch")
	cmd.execute()
End If

If Request("btnAdd.x") <> "" Then Response.Redirect "../operaciones.asp?cmd=activitySubmit&Confirm=" & Request("Confirm")
If Request("btnGeneral.x") <> "" Then Response.Redirect "../operaciones.asp?cmd=activityGeneral"
If Request("btnAddress.x") <> "" Then Response.Redirect "../operaciones.asp?cmd=activityAddress"
If Request("btnContent.x") <> "" Then Response.Redirect "../operaciones.asp?cmd=activityContent"
If Request("btnUDF.x") <> "" or Request("changeGroup") = "Y" Then Response.Redirect "../operaciones.asp?cmd=activityUDF&GroupID=" & Request("newGroupID")

Response.Redirect "../operaciones.asp?cmd=activity"



Function getActAddError()
	RetVal = ""
	
	If Request("Confirm") <> "Y" Then
		set rFlow = Server.CreateObject("ADODB.RecordSet")
		set rChk = Server.CreateObject("ADODB.RecordSet")
		
		sqlFlow = "declare @LogNum int set @LogNum = " & Session("ActRetVal") & " " & _
		"select T0.FlowID, T0.Name, Type, Query  " & _
		"from OLKUAF T0  " & _
		"inner join OLKUAF1 T1 on T1.FlowID = T0.FlowID and T1.SlpCode in (" & Session("vendid") & ",-999) " & _
		"where T0.Active = 'Y' and T0.ExecAt = 'C2' " 
		
		If Request("DocConf") <> "" Then sqlFlow = sqlFlow & " and T0.FlowID not in (" & Request("DocConf") & ") "
		
		sqlFlow = sqlFlow & " order by Type, [Order] asc"
		
		set rFlow = conn.execute(sqlFlow)
		sqlBase = 	"declare @LogNum int set @LogNum = " & Session("ActRetVal") & " " & _
					"declare @SlpCode int set @SlpCode = " & Session("VendID") & " " & _
					"declare @dbName nvarchar(100) set @dbName = db_name() " & _
					"declare @branch int set @branch = " & Session("branch") & " " & _
					"declare @CardCode nvarchar(15) set @CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "' "
		
		do while not rFlow.eof
			sql = sqlBase & rFlow("Query")
			'response.write sql
			set rChk = conn.execute(sql)
			If not rChk.eof then
				If Not IsNull(rChk(0)) Then
					If lcase(rChk(0)) = lcase("True") Then
						If ActRetVal <> "" Then ActRetVal = ActRetVal & ", "
						ActRetVal = ActRetVal & rFlow("FlowID")
						If rFlow("Type") = 0 Then Exit do
					End If
				End If
			End If
		rFlow.movenext
		loop
	End If
	getActAddError = ActRetVal
End Function

%>

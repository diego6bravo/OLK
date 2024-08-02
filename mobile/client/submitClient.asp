<%@ Language=VBScript %> 
<!--#include file="../myHTMLEncode.asp"-->
<% If session("OLKDB") = "" Then response.redirect "../lock.asp" %>
<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
  
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandText = "DBOLKCheckObjectCheckSum" & Session("ID")
cmd.CommandType = &H0004
cmd.Parameters.Refresh()
cmd("@LogNum") = Session("CrdRetVal")
cmd.Execute()

If cmd("@IsValid").Value = "N" Then
	Session("RetryRetVal") = Session("CrdRetVal")
	Session("CrdRetVal") = ""
	Response.Redirect "../operaciones.asp?cmd=crcerror"
End If

Select Case Request("cmd")
	Case "data"
		sql = 	"update r3_obscommon..TCRD set CardCode = N'" & saveHTMLDecode(Request("CardCode"), False) & "', CardName = N'" & saveHTMLDecode(Request("CardName"), False) & "', CardType = N'" & Request("CardType") & "', " & _
				"GroupCode = " & Request("GroupCode") & ", LicTradNum = N'" & saveHTMLDecode(Request("LicTradNum"), False) & "', CmpPrivate = N'" & saveHTMLDecode(Request("CmpPrivate"), False) & "', SlpCode = " & Request("SlpCode") & " " & _
				"where LogNum = " & Session("CrdRetVal")
		conn.execute(sql)
	Case "addData"
		sql = "update R3_ObsCommon..TCRD set Phone1 = N'" & saveHTMLDecode(Request("Phone1"), False) & "', Phone2 = N'" & saveHTMLDecode(Request("Phone2"), False) & "', Fax = N'" & saveHTMLDecode(Request("Fax"), False) & "', Notes = N'" & saveHTMLDecode(Request("Notes"), False) & "', " & _
							"Cellular = N'" & saveHTMLDecode(Request("Cellular"), False) & "', E_Mail = N'" & saveHTMLDecode(Request("E_Mail"), False) & "' where LogNum = " & Session("CrdRetVal")
		conn.execute(sql)
	Case "contact"
		If Request("btnSave") <> "" Then 
			If Request("Position") <> "" Then Position = "N'" & saveHTMLDecode(Request("Position"), False) & "'" Else Position = "NULL"
			If Request("Address") <> "" Then Address = "N'" & saveHTMLDecode(Request("Address"), False) & "'" Else Address = "NULL"
			If Request("Title") <> "" Then Title = "N'" & saveHTMLDecode(Request("Title"), False) & "'" Else Title = "NULL"
			If Request("Tel1") <> "" Then Tel1 = "N'" & saveHTMLDecode(Request("Tel1"), False) & "'" Else Tel1 = "NULL"
			If Request("Tel2") <> "" Then Tel2 = "N'" & saveHTMLDecode(Request("Tel2"), False) & "'" Else Tel2 = "NULL"
			If Request("Cellolar") <> "" Then Cellolar = "N'" & saveHTMLDecode(Request("Cellolar"), False) & "'" Else Cellolar = "NULL"
			If Request("Fax") <> "" Then Fax = "N'" & saveHTMLDecode(Request("Fax"), False) & "'" Else Fax = "NULL"
			If Request("EMail") <> "" Then EMail = "N'" & saveHTMLDecode(Request("EMail"), False) & "'" Else EMail = "NULL"
			If Request("Pager") <> "" Then Pager = "N'" & saveHTMLDecode(Request("Pager"), False) & "'" Else Pager = "NULL"
			If Request("Notes1") <> "" Then Notes1 = "N'" & saveHTMLDecode(Request("Notes1"), False) & "'" Else Notes1 = "NULL"
			If Request("Notes2") <> "" Then Notes2 = "N'" & saveHTMLDecode(Request("Notes2"), False) & "'" Else Notes2 = "NULL"
			If Request("Password") <> "" Then Password = "N'" & saveHTMLDecode(Request("Password"), False) & "'" Else Password = "NULL"
			If Request("BirthPlace") <> "" Then BirthPlace = "N'" & saveHTMLDecode(Request("BirthPlace"), False) & "'" Else BirthPlace = "NULL"
			If Request("BirthDate") <> "" Then BirthDate = "Convert(datetime,'" & SaveSqlDate(Request("BirthDate")) & "',120)" Else BirthDate = "NULL"
			If Request("Gender") <> "" Then Gender = "N'" & saveHTMLDecode(Request("Gender"), False) & "'" Else Gender = "NULL"
			If Request("Profession") <> "" Then Profession = "N'" & saveHTMLDecode(Request("Profession"), False) & "'" Else Profession = "NULL"

			lineNum = -1
			If Request("EditID") = "" Then
				sql = 	"declare @LogNum int set @LogNum = " & Session("CrdRetVal") & " " & _
						"declare @LineNum int set @LineNum = IsNull((select Max(LineNum)+1 from R3_ObsCommon..CRD2 where LogNum = @LogNum), 0) " & _
						"select @LineNum LineNum " & _
						"insert R3_ObsCommon..CRD2(LogNum, LineNum, LineCommand, [Name], NewName, Position, Address, Tel1, Tel2, Cellolar, Fax, E_MailL, Pager, Notes1, Notes2, Password, BirthPlace, BirthDate, Gender, Profession, Title) " & _
						"values(@LogNum, @LineNum, 'A', N'" & Request("NewName") & "', N'" & Request("NewName") & "', " & Position & ", " & Address & ", " & Tel1 & ", " & Tel2 & ", " & Cellolar & ", " & Fax & ", " & EMail & ", " & Pager & ", " & Notes1 & "," & _
						" " & Notes2 & ", " & Password & ", " & BirthPlace & ", " & BirthDate & ", " & Gender & ", " & Profession & ", " & Title & ") " & _
						"if (select CntctPrsn from R3_ObsCommon..TCRD where LogNum = @LogNum) is null or '" & Request("SetDef") & "' = 'Y' begin " & _
						"	update R3_ObsCommon..TCRD set CntctPrsn = N'" & saveHTMLDecode(Request("NewName"), False) & "' where LogNum = @LogNum " & _
						"End "
				set rs = conn.execute(sql)
				lineNum = rs("LineNum")
			Else
				lineNum = Request("EditID")
				sql = 	"declare @LogNum int set @LogNum = " & Session("CrdRetVal") & " " & _
						"declare @LineNum int set @LineNum = " & Request("EditID") & " " & _
						"if (select CntctPrsn from R3_ObsCommon..TCRD where LogNum = @LogNum) = (select Name from R3_ObsCommon..CRD2 where LogNum = @LogNum and LineNum = @LineNum) or '" & Request("SetDef") & "' = 'Y' begin " & _
						"	update R3_ObsCommon..TCRD set CntctPrsn = N'" & saveHTMLDecode(Request("NewName"), False) & "' where LogNum = @LogNum " & _
						"End " & _
						"update R3_ObsCommon..CRD2 set Name = Case LineCommand When 'A' Then N'" & saveHTMLDecode(Request("NewName"), False) & "' Else Name End, NewName = N'" & saveHTMLDecode(Request("NewName"), False) & "', Position = " & Position & ", Address = " & Address & ", Tel1 = " & Tel1 & ", Tel2  = " & Tel2 & ", Cellolar = " & Cellolar & ", " & _
						"Fax = " & Fax & ", E_MailL = " & EMail & ", Pager = " & Pager & ", Notes1 = " & Notes1 & ", Notes2 = " & Notes2 & ", Password = " & Password & ", BirthPlace = " & BirthPlace & ", BirthDate = " & BirthDate & ", " & _
						"Gender = " & Gender & ", Profession = " & Profession & ", Title = " & Title & " " & _
						"where LogNum = @LogNum and LineNum = @LineNum"
				conn.execute(sql)
			End If
			
			set rv = Server.CreateObject("ADODB.RecordSet")
			sql = "select AliasID, (select SDKID Collate database_default from r3_obscommon..tcif where companydb = N'" & Session("OLKDb") & "')+AliasID As InsertID, TypeID " & _
				  "from [" & Session("olkdb") & "]..cufd T0 " & _
				  "left outer join [" & Session("olkdb") & "]..OLKCUFD T1 on T1.TableID = T0.TableID and T1.FieldID = T0.FieldID " & _
				  "where T0.TableId = 'OCPR' and AType in ('V','T') and OP in ('T','P') and Active = 'Y'"
			set rv = server.createobject("ADODB.RecordSet")
			rv.open sql, conn, 3, 1
			
			If Not rv.Eof Then
				sql = "update R3_ObsCommon..CRD2 set "
				do while not rv.eof
					If rv.bookmark > 1 Then sql = sql & ", "
					sql = sql & rv("InsertID") & " = "
					If Request("U_" & rv("AliasID")) <> "" Then
						strVal = saveHTMLDecode(Request("U_" & rv("AliasID")), False)
						If rv("TypeID") = "D" Then sql = sql & "Convert(datetime,'" & SaveSqlDate(strVal) & "',120)" Else sql = sql & "N'" & strVal & "'"
					Else
						sql = sql & "NULL"
					End If
				rv.movenext
				loop
				sql = sql & " where LogNum = " & Session("CrdRetVal") & " and LineNum = " & lineNum
				conn.execute(sql)
			End If
		End If
	Case "address"
		If Request("btnSave") <> "" Then
			defFld = ""
			Select Case Request("AdresType")
				Case "S"
					defFld = "ShipToDef"
				Case "B"
					defFld = "BillToDef"
			End Select
			
			If Request("Street") <> "" Then Street = "N'" & saveHTMLDecode(Request("Street"), False) & "'" Else Street = "NULL"
			If Request("Block") <> "" Then Block = "N'" & saveHTMLDecode(Request("Block"), False) & "'" Else Block= "NULL"
			If Request("City") <> "" Then City = "N'" & saveHTMLDecode(Request("City"), False) & "'" Else City= "NULL"
			If Request("ZipCode") <> "" Then ZipCode = "N'" & saveHTMLDecode(Request("ZipCode"), False) & "'" Else ZipCode = "NULL"
			If Request("County") <> "" Then County = "N'" & saveHTMLDecode(Request("County"), False) & "'" Else County = "NULL"
			If Request("Country") <> "" Then Country = "N'" & saveHTMLDecode(Request("Country"), False) & "'" Else Country = "NULL"
			If Request("State") <> "" Then State = "N'" & saveHTMLDecode(Request("State"), False) & "'" Else State = "NULL"
			If Request("TaxCode") <> "" Then TaxCode = "N'" & saveHTMLDecode(Request("TaxCode"), False) & "'" Else TaxCode = "NULL"
			
			lineNum = -1
			If Request("EditID") = "" Then
				sql = 	"declare @LogNum int set @LogNum = " & Session("CrdRetVal") & " " & _
						"declare @LineNum int set @LineNum = IsNull((select Max(LineNum)+1 from R3_ObsCommon..CRD1 where LogNum = @LogNum), 0) " & _
						"select @LineNum LineNum " & _
						"insert R3_ObsCommon..CRD1(LogNum, LineNum, LineCommand, Address, NewAddress, Street, AdresType, Block, City, ZipCode, County, Country, State, TaxCode) " & _
						"values(@LogNum, @LineNum, 'A', N'" & saveHTMLDecode(Request("NewAddress"), False) & "', N'" & saveHTMLDecode(Request("NewAddress"), False) & "', " & Street & ", '" & Request("AdresType") & "', " & Block & ", " & City & ", " & ZipCode & ", " & County & ", " & Country & ", " & State & ", " & TaxCode & ") " & _
						"if (select " & defFld & " from R3_ObsCommon..TCRD where LogNum = @LogNum) is null or '" & Request("SetDef") & "' = 'Y' begin " & _
						"	update R3_ObsCommon..TCRD set " & defFld & " = N'" & saveHTMLDecode(Request("NewAddress"), False) & "' where LogNum = @LogNum " & _
						"End "
				set rs = conn.execute(sql)
				lineNum = rs("LineNum")
			Else
				lineNum = Request("EditID")
				sql = 	"declare @LogNum int set @LogNum = " & Session("CrdRetVal") & " " & _
						"declare @LineNum int set @LineNum = " & Request("EditID") & " " & _
						"If (select " & defFld & " from R3_ObsCommon..TCRD where LogNum = @LogNum) = (select NewAddress from R3_ObsCommon..CRD1 where LogNum = @LogNum and LineNum = @LineNum) or '" & Request("SetDef") & "' = 'Y' begin " & _
						"	update R3_ObsCommon..TCRD set " & defFld & " = N'" & saveHTMLDecode(Request("NewAddress"), False) & "' where LogNum = @LogNum " & _
						"End " & _
						"update R3_ObsCommon..CRD1 set Address = Case LineCommand When 'A' Then N'" & saveHTMLDecode(Request("NewAddress"), False) & "' Else Address End, NewAddress = N'" & saveHTMLDecode(Request("NewAddress"), False) & "', " & _
						"Street = " & Street & ", Block = " & Block & ", City = " & City & ", ZipCode = " & ZipCode & ", County = " & County & ", Country = " & Country & ", State = " & State & ", TaxCode = " & TaxCode & " " & _
						"where LogNum = @LogNum and LineNum = @LineNum"
				conn.execute(sql)
			End If
			
			set rv = Server.CreateObject("ADODB.RecordSet")
			sql = "select AliasID, (select SDKID Collate database_default from r3_obscommon..tcif where companydb = N'" & Session("OLKDb") & "')+AliasID As InsertID, TypeID " & _
				  "from [" & Session("olkdb") & "]..cufd T0 " & _
				  "left outer join [" & Session("olkdb") & "]..OLKCUFD T1 on T1.TableID = T0.TableID and T1.FieldID = T0.FieldID " & _
				  "where T0.TableId = 'CRD1' and AType in ('V','T') and OP in ('T','P') and Active = 'Y'"
			set rv = server.createobject("ADODB.RecordSet")
			rv.open sql, conn, 3, 1
			
			If Not rv.Eof Then
				sql = "update R3_ObsCommon..CRD1 set "
				do while not rv.eof
					If rv.bookmark > 1 Then sql = sql & ", "
					sql = sql & rv("InsertID") & " = "
					If Request("U_" & rv("AliasID")) <> "" Then
						strVal = saveHTMLDecode(Request("U_" & rv("AliasID")), False)
						If rv("TypeID") = "D" Then sql = sql & "Convert(datetime,'" & SaveSqlDate(strVal) & "',120)" Else sql = sql & " N'" & strVal & "'"
					Else
						sql = sql & "NULL"
					End If
				rv.movenext
				loop
				sql = sql & " where LogNum = " & Session("CrdRetVal") & " and LineNum = " & lineNum
				conn.execute(sql)
			End If
		End If
	Case "UDF"
		set rs = Server.CreateObject("ADODB.RecordSet")
		sql = "select AliasID, (select SDKID Collate database_default from r3_obscommon..tcif where companydb = N'" & Session("OLKDb") & "')++AliasID As InsertID, TypeID " & _
			  "from cufd T0 " & _
			  "left outer join OLKCUFD T1 on T1.TableID = T0.TableID and T1.FieldID = T0.FieldID " & _
			  "where T0.TableId = 'OCRD' and AType in ('V','T') and OP in ('T','P') and Active = 'Y' and GroupID = " & Request("GroupID")
		set rs = server.createobject("ADODB.RecordSet")
		rs.open sql, conn, 3, 1
		
		sql = "update R3_ObsCommon..TCRD set "
		do while not rs.eof
			If rs.bookmark > 1 Then sql = sql & ", "
			sql = sql & rs("InsertID") & " = "
			If Request("U_" & rs("AliasID")) <> "" Then
				strVal = saveHTMLDecode(Request("U_" & rs("AliasID")), False)
				If rs("TypeID") = "D" Then sql = sql & "Convert(datetime,'" & SaveSqlDate(strVal) & "',120)" Else sql = sql & "N'" & strVal & "'"
			Else
				sql = sql & "NULL"
			End If
		rs.movenext
		loop
		sql = sql & " where LogNum = " & Session("CrdRetVal")
		conn.execute(sql)
End Select
If Request("btnUpdate.x") <> "" or Request("btnAdd.x") <> "" Then
	AddErr = getCrdAddError()
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

If Request("btnUpdate.x") <> "" Then TransType = "U" Else TransType = "A"

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKPostObjectCreation" & Session("ID")
cmd.Parameters.Refresh
cmd("@sessiontype") = "P"
cmd("@transtype") = TransType
cmd("@object") = 2 
cmd("@LogNum") = Session("CrdRetVal")
cmd("@CurrentSlpCode") = Session("vendid")
cmd("@Branch") = Session("branch")
cmd.execute()

conn.close
If Request("btnGeneral.x") <> "" Then
	response.redirect "../operaciones.asp?cmd=newClientAddData"
ElseIf Request("btnUDF.x") <> "" or Request("changeGroup") = "Y" Then 
	response.redirect "../operaciones.asp?cmd=newClientUDF&GroupID=" & Request("newGroupID")
ElseIf Request("btnAddress.x") <> "" or Request("cmd") = "address" and Request("btnSave") <> "" Then
		response.redirect "../operaciones.asp?cmd=newClientAddresses"
ElseIf Request("btnContacts.x") <> "" or Request("cmd") = "contact" and Request("btnSave") <> "" Then
		response.redirect "../operaciones.asp?cmd=newClientContacts"
ElseIf Request("btnAdd.x") <> "" or Request("btnUpdate.x") <> "" Then
		response.redirect "../operaciones.asp?cmd=newClientSubmit&Confirm=" & Request("Confirm")
Else
		response.redirect "../operaciones.asp?cmd=newClient"
End If


Function getCrdAddError()
	RetVal = ""
	
	If Request("Confirm") <> "Y" Then
		set rFlow = Server.CreateObject("ADODB.RecordSet")
		set rChk = Server.CreateObject("ADODB.RecordSet")
		
		sqlFlow = "declare @LogNum int set @LogNum = " & Session("CrdRetVal") & " " & _
		"select T0.FlowID, T0.Name, Type, Query  " & _
		"from OLKUAF T0  " & _
		"inner join OLKUAF1 T1 on T1.FlowID = T0.FlowID and T1.SlpCode in (" & Session("vendid") & ",-999) " & _
		"where T0.Active = 'Y' and T0.ExecAt = 'C1' " 
		
		If Request("DocConf") <> "" Then sqlFlow = sqlFlow & " and T0.FlowID not in (" & Request("DocConf") & ") "
		
		sqlFlow = sqlFlow & " order by Type, [Order] asc"
		
		set rFlow = conn.execute(sqlFlow)
		sqlBase = 	"declare @LogNum int set @LogNum = " & Session("CrdRetVal") & " " & _
					"declare @SlpCode int set @SlpCode = " & Session("VendID") & " " & _
					"declare @dbName nvarchar(100) set @dbName = db_name() " & _
					"declare @branch int set @branch = " & Session("branch") & " "
		
		do while not rFlow.eof
			sql = sqlBase & rFlow("Query")
			'response.write sql
			set rChk = conn.execute(sql)
			If not rChk.eof then
				If Not IsNull(rChk(0)) Then
					If lcase(rChk(0)) = lcase("True") Then
						If RetVal <> "" Then RetVal = RetVal & ", "
						RetVal = RetVal & rFlow("FlowID")
						If rFlow("Type") = 0 Then Exit do
					End If
				End If
			End If
		rFlow.movenext
		loop
	End If
	getCrdAddError = RetVal
End Function

%>

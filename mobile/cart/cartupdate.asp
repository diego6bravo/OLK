<%@ Language=VBScript %> 
<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

%>
<!--#include file="../myHTMLEncode.asp"-->
<!--#include file="../lcidReturn.inc"-->
<!--#include file="../authorizationClass.asp"-->
<% If session("OLKDB") = "" Then response.redirect "../lock.asp" %>
<%
dim varx
dim varxx
Dim varpricex

Dim myAut
set myAut = New clsAuthorization

varx = "0"
           set rd = Server.CreateObject("ADODB.recordset")
           set rs = Server.CreateObject("ADODB.recordset")

sql = "select Object from R3_ObsCommon..TLOG where LogNum = " & Session("RetVal") 
set rs = conn.execute(sql)
myLinesCount = getLinesCount
Object = CInt(rs("Object"))
	
sqlnum = "select LineNum, T0.ItemCode, WhsCode from r3_obscommon..doc1 T0 " & _
"where LogNum = " & Session("RetVal")

If myApp.EnableCartSum and Request("oldViewMode") = "" Then
	If myApp.CartSumQty < myLinesCount and Request("document") <> "B" Then
		sqlnum = sqlnum & " and T0.LineNum >= (select Min(LineNum) from (select top " & myApp.CartSumQty & " LineNum from R3_ObsCommon..DOC1 X0 where LogNum = " & Session("RetVal") & " order by LineNum desc) T0)"
	Else
		If Request("String") <> "" Then
			arrSearchStr = Split(Request("String"), " ")
			sqlSearchFilter = ""
			For i = 0 to UBound(arrSearchStr)
				If sqlSearchFilter <> "" Then sqlSearchFilter = sqlSearchFilter & " or "
				sqlSearchFilter = sqlSearchFilter & " (ItemCode like N'%" & arrSearchStr(i) & "%' or ItemName like N'%" & arrSearchStr(i) & "%'" & _
	  											" or frgnName like N'%" & arrSearchStr(i) & "%') "
			Next
			sqlnum = sqlnum & " and T0.ItemCode in (select ItemCode collate database_default from OITM where ((" & sqlSearchFilter & ") or CodeBars = N'" & Request("String") & "')) "
		End If
	End If
End If

If Request("btnDel.x") <> "" and Request("chkDel") <> "" Then sqlnum = sqlnum & " and T0.LineNum not in (" & Request("chkDel") & ")"

set rd = conn.execute(sqlnum)

If Request("Draft") = "Y" Then
	sql = "update R3_ObsCommon..TLOG set Draft = 'Y' where LogNum = " & Session("RetVal")
	conn.execute(sql)
End If

If Request("Authorize") = "Y" and Object = 17 Then
	sql = "update R3_ObsCommon..TDOC set Confirmed = 'N' where LogNum = " & Session("RetVal")
	conn.execute(sql)
End If

sqlx = "update r3_obscommon..tdoc set cntctcode = '" & request.form("CntctCode") & "', " & _
	   "Series = IsNull(" & myAut.GetObjectProperty(Object, "S") & ", (select Series from olkDocConf where ObjectCode = " & Object & ")), CardName = N'" & saveHTMLDecode(Request("CardName"), False) & "', DiscPrcnt = " & getNumeric(Request("DiscPrcnt")) & " where lognum = " & Session("retval")
conn.execute(sqlx)

varErr = ""
do while not rd.eof
varx = rd("LineNum")
varxx = "T" & varx
varpricex = "Price" & varx
if Rd("WhsCode") <> "" Then WhsCode = "'" & Rd("WhsCode") & "'" Else WhsCode = "NULL"
If Object <> 23 Then
	sql = "select OLKCommon.dbo.DBOLKItemInv" & Session("ID") & "('" & Rd("ItemCode") & "', " & WhsCode & ", " & getNumeric(Request(varxx)) & ", '" & Session("olkdb") & "', " & Session("RetVal") & ", " & varx & ") Verfy"
	set rc = conn.execute(sql)
	Verfy = rc("Verfy") = "Y"
Else
	Verfy = True
End If
	If Verfy Then
		sql = "exec OLKCommon..DBOLKCartUpdateLines" & Session("ID") & " " & Session("RetVal") & ", " & RD("LineNum") & ", '" & getNumeric(Request(varpricex)) & "', " & getNumeric(Request(varxx)) & ", " & Request("un" & RD("LineNum")) & ", null, null"
		conn.execute(sql)
		
		If Request("ManPrc" & rd("LineNum")) = "Y" Then ManPrc = "Y" Else ManPrc = "N"
		
		sql = "update OLKSalesLines set ManPrc = '" & ManPrc & "' where LogNum = " & Session("RetVal") & " and LineNum = " & rd("LineNum")
		conn.execute(sql)
	Else
		varErr = "&" & varx & "=Y"
	End If
rd.MoveNext
loop

If Request("btnDel.x") <> "" and Request("chkDel") <> "" Then
	sql = "delete R3_ObsCommon..doc1 where lognum = " & Session("RetVal") & " and LineNum in (" & Request("chkDel") & ") " & _
	" delete olkSalesLines where LogNum = " & Session("RetVal") & " and LineNum in (" & Request("chkDel") & ") "
	conn.execute(sql)
End If



If Request.Form("I1.x") <> "" or Request("btnViewAll") <> "" or varErr <> "" or Request("btnDel.x") <> "" Then
	ViewMode = Request("ViewMode")
		If Request("btnViewAll") <> "" Then
			If ViewMode = "" Then ViewMode = "all" Else ViewMode = ""
		End If
		
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKPostObjectCreation" & Session("ID")
	cmd.Parameters.Refresh
	cmd("@sessiontype") = "P"
	cmd("@transtype") = "U"
	cmd("@object") = Object 
	cmd("@LogNum") = Session("RetVal")
	cmd("@CurrentSlpCode") = Session("vendid")
	cmd("@Branch") = Session("branch")
	cmd.execute()
	response.redirect "../operaciones.asp?cmd=cart" & "&ViewMode=" & ViewMode & "&document=" & Request("document") & "&string=" & Request("string") & varErr
ElseIf Request.Form("I2.x") <> "" Then
	AddErr = getDocAddError()
	If AddErr <> "" Then
		retURL = ""
		For each itm in Request.Form
			If LCase(itm) <> "item" and itm <> "Confirm" Then
				If retURL <> "" Then retURL = retURL & "{a}"
				retURL = retURL & itm & "{e}" & Server.URLEncode(Request(itm))
			End If
		Next
		For each itm in Request.QueryString
			If LCase(itm) <> "item" and itm <> "Confirm" Then
				If retURL <> "" Then retURL = retURL & "{a}"
				retURL = retURL & itm & "{e}" & Server.URLEncode(Request(itm))
			End If
		Next
		response.redirect "../operaciones.asp?cmd=DocFlowErr&DocFlowErr=" & AddErr & "&retURL=" & retURL
	End If
	
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKValidateObjectFields" & Session("ID")
	cmd.Parameters.Refresh
	cmd("@LogNum") = Session("RetVal")
	cmd("@UserType") = userType
	cmd("@LanID") = Session("LanID")
	set rv = Server.CreateObject("ADODB.Recordset")
	set rv = cmd.execute()
	If Not rv.Eof Then
		Response.Redirect "../operaciones.asp?cmd=cartopt&compFld=Y"
	End If
	
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKPostObjectCreation" & Session("ID")
	cmd.Parameters.Refresh
	cmd("@sessiontype") = "P"
	cmd("@transtype") = "A"
	cmd("@object") = Object 
	cmd("@LogNum") = Session("RetVal")
	cmd("@CurrentSlpCode") = Session("vendid")
	cmd("@Branch") = Session("branch")
	cmd.execute()
	response.redirect "../operaciones.asp?cmd=cartSubmit&Confirm=" & Request("Confirm")
End If

conn.close 
set rs = nothing
set rd = nothing


Function getLinesCount()
	sql = "SELECT Count('A') from R3_OBSCommon..DOC1 T0  " & _
	"inner join OITM T1 on T1.ItemCode = T0.ItemCode collate database_default  " & _
	"Where T0.LogNum = " & Session("RetVal") 
	set rCount = Server.CreateObject("ADODB.RecordSet")
	set rCount = conn.execute(sql)
	getLinesCount = rCount(0)
End Function

Function getDocAddError()
	RetVal = ""
	
	If Request("Confirm") <> "Y" Then
		set rFlow = Server.CreateObject("ADODB.RecordSet")
		set rChk = Server.CreateObject("ADODB.RecordSet")
		
		sqlFlow = "declare @LogNum int set @LogNum = " & Session("RetVal") & " " & _
		"declare @ObjectCode int  " & _
		"If Exists(select 'A' from olkcic where InvLogNum = @LogNum and PayLogNum is not null) begin " & _
		"	set @ObjectCode = 48 " & _
		"End Else If (select Object from R3_ObsCommon..TLOG where LogNum = @LogNum) = 13 and (select ReserveInvoice from R3_ObsCommon..TDOC where LogNum = @LogNum) = 'Y' Begin " & _
		"	set @ObjectCode = -13 " & _
		"end else begin " & _
		"	set @ObjectCode = (select Object from R3_ObsCommon..TLOG where LogNum = @LogNum) " & _
		"end " & _
		"select T0.FlowID, T0.Name, Type, Query  " & _
		"from OLKUAF T0  "
		
		If userType = "V" Then
			sqlFlow = sqlFlow & "inner join OLKUAF1 T1 on T1.FlowID = T0.FlowID and T1.SlpCode in (" & Session("vendid") & ",-999) "
		End If
		
		sqlFlow = sqlFlow & "inner join OLKUAF2 T2 on T2.FlowID = T0.FlowID " & _
		"where T2.ObjectCode = @ObjectCode and T0.Active = 'Y' and T0.ExecAt = 'D3' "
		
		If userType = "C" Then
			sqlFlow = sqlFlow & " and T0.ApplyToClient = 'Y' "
		End If
		
		If Request("DocConf") <> "" Then sqlFlow = sqlFlow & " and T0.FlowID not in (" & Request("DocConf") & ") "
		
		sqlFlow = sqlFlow & " order by Type, [Order] asc"
		'response.redirect "http://www.topmanage.com.pa/query.asp?query=" & sqlFlow
		
		set rFlow = conn.execute(sqlFlow)
		sqlBase = 	"declare @LogNum int set @LogNum = " & Session("RetVal") & " " & _
					"declare @CardCode nvarchar(15) set @CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "' " & _
					"declare @SlpCode int set @SlpCode = " & Session("VendID") & " " & _
					"declare @dbName nvarchar(100) set @dbName = db_name() " & _
					"declare @branch int set @branch = " & Session("branch") & " "
		
		do while not rFlow.eof
			sql = sqlBase & rFlow("Query")
			sql = QueryFunctions(sql)
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
	getDocAddError = RetVal
	'response.redirect "http://www.topmanage.com.pa/query.asp?query=" & RetVal
End Function
%>


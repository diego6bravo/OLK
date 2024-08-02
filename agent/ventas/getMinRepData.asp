<!--#include file="../myHTMLENcode.asp"-->
<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="../chkLogin.asp" -->
<% 
set rs = Server.CreateObject("ADODB.RecordSet")


set rx = Server.CreateObject("ADODB.RecordSet")
sql = 	"select T0.RowType+Convert(nvarchar(20),T0.LineIndex) RowID, T0.RowQuery " & _
		"from OLKCMREP T0 " & _
		"where T0.RowActive = 'Y' and T0.ShowV = 'Y' and RowQuery is not null " & _
		"order by T0.RowOrder asc"
rx.open sql, conn, 3, 1


if not rx.eof then
	sql = "declare @LogNum int set @LogNum = " & Session("RetVal") & " " & _
	"declare @CardCode nvarchar(15) set @CardCode = N'" & saveHTMLDecode(Session("username"), False) & "' " & _
	"declare @LanID int set @LanID = " & Session("LanID") & " " & _
	"select "
	do while not rx.eof
		If rx.bookmark > 1 Then sql = sql & ", "
		sql = sql & "(" & rx("RowQuery") & ") As '" & rx("RowID") & "'"
	rx.movenext
	loop
	sql = QueryFunctions(sql)
	set rx = conn.execute(sql)
	
	strRetVal = ""
	For each fld in rx.Fields
		If strRetVal <> "" Then strRetVal = strRetVal & "{S}"
		strRetVal = strRetVal & fld.Name & "{=}" & fld
	Next
End If

Response.Write strRetVal %>

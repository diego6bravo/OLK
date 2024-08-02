<!--#include file="../chkLogin.asp" -->
<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

Server.ScriptTimeout = 100000

sql = "select dbName from OLKCommon..OLKDBA where ID = " & Request("dbName")
set rs = conn.execute(sql)
dbName = rs(0)
dbID = Request("dbName")

conn.execute("use [" & dbName & "]")

set rs = Server.CreateObject("ADODB.RecordSet")
sql = "select TableID, Type from olkcommon..olktdb where RTrim(TableID) <> ''"
set rs = conn.execute(sql)
do while not rs.eof
	Select Case rs("Type")
		Case "T"
			sql = "if exists (select '' from dbo.sysobjects where id = object_id(N'[dbo].[" & rs("TableID") & "]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & _
				  "drop table [dbo].[" & rs("TableID") & "]"
			conn.execute(sql)	
		Case "P"
			sql = "if exists (select '' from sysobjects where id = object_id(N'[dbo].[DB" & rs("TableID") & dbID & "]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & _
				  "drop procedure [DB" & rs("TableID") & dbID & "]"
			connCommon.execute(sql)	
		Case "F"
			sql = "if exists (select '' from sysobjects where id = object_id(N'[dbo].[DB" & rs("TableID") & dbID & "]') and xtype in (N'FN', N'IF', N'TF'))" & _
				  "drop function [DB" & rs("TableID") & dbID & "]"
			connCommon.execute(sql)	
	End Select

rs.movenext
loop
sql = "delete olkcommon..OLKDBA where dbName = '" & dbName & "' " & _
"delete TMMailer..TMMailLines where mailID in (select mailID from TMMailer..TMMail where mailID < 0 and dbName = '" & dbName & "') " & _
"delete TMMailer..TMMail where mailID < 0 and dbName = '" & dbName & "'"
conn.execute(sql)
conn.close
Session("OlkDB") = ""
Session("ID") = ""
response.redirect "../admin.asp?cmd=home" 

%>
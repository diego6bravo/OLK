<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

Dim oRs
Dim sSQL
Dim nFileID

nFileID = Request.QueryString("FileID")

If Not nFileID = "" And IsNumeric(nFileID) Then

	Set oRs = Server.CreateObject("ADODB.Recordset")

	sSQL = "SELECT FileName, ContentType, BinaryData FROM OLKImgFiles WHERE FileID = " & Request.QueryString("FileID")

	oRs.Open sSQL, conn, 3, 3

	If Not oRs.EOF Then
		Response.ContentType = oRs(1)
		Response.BinaryWrite oRs(2)
	Else
		Response.Write("File could not be found")
	End If

	oRs.Close

	Set oRs = Nothing
Else
	Response.Write("File could not be found")
End If
%>


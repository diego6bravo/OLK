<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<!--#INCLUDE FILE="clsUpload.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

Dim objUpload
Dim strFileName
Dim objRs
Dim lngFileID

' Instantiate Upload Class
Set objUpload = New clsUpload

' Grab the file name
strFileName = objUpload.Fields("File1").FileName

Set objRs = Server.CreateObject("ADODB.Recordset")


objRs.Open "OLKImgFiles", conn, 3, 3

objRs.AddNew

objRs.Fields("FileName").Value = objUpload.Fields("File1").FileName
objRs.Fields("FileSize").Value = objUpload.Fields("File1").Length
objRs.Fields("ContentType").Value = objUpload.Fields("File1").ContentType
objRs.Fields("BinaryData").AppendChunk objUpload("File1").BLOB & ChrB(0)

objRs.Update

objRs.Close

objRs.Open "SELECT Max(FileID) AS ID FROM OLKImgFiles", conn, 3, 3
lngFileID = objRs.Fields("ID").Value
objRs.Close

Set objRs = Nothing
Set conn = Nothing
Set objUpload = Nothing
%>
<SCRIPT LANGUAGE="JavaScript">
opener.changepic('<%=lngFileID%>');
window.close()
</SCRIPT>

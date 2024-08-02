<!--#include file="../myHTMLEncode.asp"-->
<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="../chkLogin.asp" -->
<!--#include file="adminTradSave.asp"-->
<% 
set rs = Server.CreateObject("ADODB.recordset")

Select Case Request("submitCmd")
	Case "del"
		sql = "update olkNews set Status = 'D' where newsIndex in (" & Request("delID") & ")"
		conn.execute(sql)
		conn.close
	Case "delNews"
		sql = "update olkNews set Status = 'D' where newsIndex = " & Request("newsIndex")
		conn.execute(sql)
		conn.close
	Case "updateNews"
	    updateNews
	Case "addNews"
		addNews
End Select

Private Sub updateNews()
    If Request("newsSource") <> "" Then newsSource = "N'" & saveHTMLDecode(Request("newsSource"), False) & "'" else newsSource = "NULL"
    If Request("newsImg") <> "" Then newsImg = "N'" & Request("newsImg") & "'" Else newsImg = "NULL"
    If Request("chkActive") = "A" Then chkActive = "A" Else chkActive = "N"
    
	varx = Request.ServerVariables("URL")
	varText = saveHTMLDecode(Request("newsText"), False)
	varText = Replace(varText,GetHTTPStr & Request.ServerVariables("SERVER_NAME") & Mid(varx,1,Len(varx)-21),"../")

	sql = "update olkNews set newsTitle = N'" & saveHTMLDecode(Request("newsTitle"), False) & "', newsDate = Convert(datetime,'" & SaveSqlDate(Request("newsDate")) & "',120), " & _
		  "newsSmallText = N'" & Left(saveHTMLDecode(Request("newsSmallText"), False),155) & "', newsText = N'" & varText & "', newsSource = " & newsSource & ", " & _
		  "newsImg = " & newsImg & ", Status = '" & chkActive & "' where newsIndex = " & Request("newsIndex")
	conn.execute(sql)
	conn.close 
	
	If Request("btnSave") <> "" Then
		response.redirect "../newsman.asp"
	Else
		response.redirect "../newsEdit.asp?newsIndex=" & Request("newsIndex")
	End If
End Sub

Private Sub addNews()
	set rs = server.createobject("ADODB.RecordSet")
	
	sql = "select ISNULL(max(newsIndex)+1,1) newsIndex from olkNews"
	set rs = conn.execute(sql)
	newsIndex = rs("newsIndex")
	set rs = nothing
	
    If Request("newsSource") <> "" Then newsSource = "N'" & saveHTMLDecode(Request("newsSource"), False) & "'" else newsSource = "NULL"
    If Request("newsImg") <> "" Then newsImg = "N'" & Request("newsImg") & "'" Else newsImg = "NULL"
    If Request("chkActive") = "A" Then chkActive = "A" Else chkActive = "N"

	varx = Request.ServerVariables("URL")
	varText = saveHTMLDecode(Request("newsText"), False)
	varText = Replace(varText,GetHTTPStr & Request.ServerVariables("SERVER_NAME") & Mid(varx,1,Len(varx)-21),"../")

	sql = "insert olkNews(newsIndex, newsTitle, newsDate, newsSmallText, newsText, newsSource, newsImg, Status) " & _
		  "values(" & newsIndex & ", N'" & saveHTMLDecode(Request("newsTitle"), False) & "', Convert(datetime,'" & SaveSqlDate(Request("newsDate")) & "',120), " & _
		  "N'" & Left(saveHTMLDecode(Request("newsSmallText"), False),155) & "', N'" & varText & "', " & newsSource & ",  " & newsImg & ", '" & chkActive & "')"
	conn.execute(sql)
	
	If Request("newsTitleTrad") <> "" Then
		SaveNewTrad Request("newsTitleTrad"), "News", "newsIndex", "alterNewsTitle", newsIndex
	End If
	
	If Request("newsSourceTrad") <> "" Then
		SaveNewTrad Request("newsSourceTrad"), "News", "newsIndex", "alterNewsSource", newsIndex
	End If

	If Request("newsSmallTextTrad") <> "" Then
		SaveNewTrad Request("newsSmallTextTrad"), "News", "newsIndex", "alterNewsSmallText", newsIndex
	End If
	
	If Request("newsTextTrad") <> "" Then
		SaveNewTrad Request("newsTextTrad"), "News", "newsIndex", "alterNewsText", newsIndex
	End If
	
	conn.close 
	
	If Request("btnSave") <> "" Then
		response.redirect "../newsman.asp"
	Else
		response.redirect "../newsEdit.asp?newsIndex=" & newsIndex
	End If
End Sub


Response.Redirect "../newsman.asp"
 %>
<!--#include file="clientInc.asp"-->
<% Select Case userType
Case "C" %><!--#include file="clientTop.asp"-->
<% Case "V" %><!--#include file="agentTop.asp"-->
<% End Select %>
<% addLngPathStr = "" %>
<!--#include file="lang/rnews.asp" -->
<%
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetObjectData" & Session("ID")
cmd.Parameters.Refresh()
cmd("@ObjType") = "S"
cmd("@ObjID") = 3
cmd("@UserType") = userType
set ra = cmd.execute()
strContent = ra("ObjContent")
If LtxtPubDat <> "" Then strContent = Replace(strContent, "{txtPubDat}", LtxtPubDat)
If LtxtBackTo <> "" Then strContent = Replace(strContent, "{txtBackTo}", LtxtBackTo)

set rn = server.CreateObject("ADODB.RecordSet")
set cmd = Server.CreateObject("ADODB.Command")
cmd.CommandText = "DBOLKGetNewsDetails" & Session("ID")
cmd.CommandType = adCmdStoredProc 
cmd.ActiveConnection = connCommon
cmd.Parameters.Refresh
cmd("@LanID") = Session("LanID")
cmd("@newsIndex") = Request("newsIndex")
set rn = cmd.execute 

If rn("Status") = "A" Then

If rn("newsImg") <> "" Then
	Pic = rn("newsImg")
Else
	Pic = "n_a.gif"
End If

strContent = Replace(strContent, "{txtPubDat}", getrnewsLngStr("LtxtPubDat"))
strContent = Replace(strContent, "{txtBackTo}", getrnewsLngStr("LtxtBackTo"))

strContent = Replace(strContent, "{NewsTblTitle}", txtNewss)
strContent = Replace(strContent, "{NewTitle}", txtNews)
strContent = Replace(strContent, "{NewsImg}", Pic)
strContent = Replace(strContent, "{ImgMaxSize}", ra("ImgMaxSize"))
strContent = Replace(strContent, "{dbName}", Session("olkdb"))
strContent = Replace(strContent, "{SelDes}", SelDes)
strContent = Replace(strContent, "{rtl}", Session("rtl"))
strContent = Replace(strContent, "{NewsTitle}", rn("newsTitle"))
strContent = Replace(strContent, "{NewsDate}", FormatDate(rn("newsDate"), True))
If Not IsNull(rn("newsSource")) Then
	strContent = Replace(strContent, "{NewsSource}", rn("newsSource") & " - ")
Else
	strContent = Replace(strContent, "{NewsSource}", "")
End If
strContent = Replace(strContent, "{NewsText}", Replace(Replace(rn("newsText"),VbCrLf,"<br>"), "<A href", "<A class=""LinkTop"" href"))
strContent = strContent %>
<%=strContent%>
<% Else %>
<script>window.history.go(-1);</script>
<% End If %>
<% rn.close %> 
<% set rn = nothing %>
<% Select Case userType
Case "C" %><!--#include file="clientBottom.asp"-->
<% Case "V" %><!--#include file="agentBottom.asp"-->
<% End Select %>
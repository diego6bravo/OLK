<!--#include file="clientInc.asp"-->
<% Select Case userType
Case "C" %><!--#include file="clientTop.asp"-->
<% Case "V" %><!--#include file="agentTop.asp"-->
<% End Select %>
<% If Not optNews and userType = "C" Then Response.Redirect "defualt.asp" %>
<% addLngPathStr = "" %>
<!--#include file="lang/news.asp" -->
<%
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetObjectData" & Session("ID")
cmd.Parameters.Refresh()
cmd("@ObjType") = "S"
cmd("@ObjID") = 2
cmd("@UserType") = userType
set ra = cmd.execute()
strContent = ra("ObjContent")
If getnewsLngStr("DtxtMore") <> "" Then strContent = Replace(strContent, "{txtMore}", getnewsLngStr("DtxtMore"))

set rn = server.CreateObject("ADODB.RecordSet")
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetTopXNews" & Session("ID")
cmd.Parameters.Refresh()
cmd("@LanID") = Session("LanID")
cmd("@TopX") = ra("TopX")
cmd("@MaxShortTextLen") = ra("MaxShortTextLen")
set rn = cmd.execute()

strContent = Replace(strContent, "{NewsTblTitle}", txtNewss)
strContent = Replace(strContent, "{SelDes}", SelDes)
strContent = Replace(strContent, "{rtl}", Session("rtl"))
strContent = Replace(strContent, "{ImgMaxSize}", ra("ImgMaxSize"))
strContent = Replace(strContent, "{dbName}", Session("olkdb"))
If Session("rtl") = "" Then
	strContent = Replace(strContent, "{txtMoreAlign}", "right")
Else
	strContent = Replace(strContent, "{txtMoreAlign}", "left")
End If

strHdr = Left(strContent, InStr(strContent, "<!--starttoploop-->")-1)
strMdl = Mid(strContent, InStr(strContent, "<!--endtopsep-->")+16, InStr(strContent, "<!--startbottom-->")-InStr(strContent, "<!--endtopsep-->")-16)
strCnt = Mid(strContent, InStr(strContent, "<!--starttoploop-->")+19, InStr(strContent, "<!--endtoploop-->")-InStr(strContent, "<!--starttoploop-->")-19)
strSep = Mid(strContent, InStr(strContent, "<!--starttopsep-->")+18, InStr(strContent, "<!--endtopsep-->")-InStr(strContent, "<!--starttopsep-->")-18)
strAltCnt = Mid(strContent, InStr(strContent, "<!--startbottom-->")+18, InStr(strContent, "<!--endstartbottom-->")-InStr(strContent, "<!--startbottom-->")-18)
strAltCntNews = Mid(strAltCnt, InStr(strAltCnt, "<!--startbottomloop-->")+24, InStr(strAltCnt, "<!--endbottomloop-->")-InStr(strAltCnt, "<!--startbottomloop-->")-24)
strBtm = Right(strContent, Len(strContent)-InStr(strContent, "<!--endstartbottom-->")-21)

%>
<%=strHdr%>
<% do while not rn.eof %>
<%=getLineCnt(strCnt)%>
<% rn.movenext %>
<% If Not rn.Eof Then Response.Write strSep %>
<% loop %>
<%=strMdl%>
<%
rn.close
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetTopXNewsNext" & Session("ID")
cmd.Parameters.Refresh()
cmd("@TopX") = ra("TopX")
cmd("@LanID") = Session("LanID")
cmd("@MaxShortTextLen") = ra("MaxShortTextLen")
set rn = cmd.execute()
If Not rn.Eof Then
	Response.Write Left(strAltCnt, InStr(strAltCnt, "<!--startbottomloop-->")-1)
%>
<% do while not rn.eof %>
<%=getLineCnt(strAltCntNews)%>
<% rn.movenext
loop %>
<% Response.Write Right(strAltCnt, Len(strAltCnt) - InStr(strAltCnt, "<!--endbottomloop-->")-21)
End If %>
<%=strBtm%><% Function getLineCnt(str)
	retVal = str
	retVal = Replace(retVal, "{newsDate}", FormatDate(rn("newsDate"), True))
	retVal = Replace(retVal, "{newsSmallText}", rn("newsSmallText"))
	retVal = Replace(retVal, "{newsIndex}", rn("newsIndex"))
	If Not IsNull(rn("newsImg")) Then retVal = Replace(retVal, "{newsPic}", rn("newsImg")) Else retVal = Replace(retVal, "{newsPic}", "n_a.gif")
	retVal = Replace(retVal, "{NewsTitle}", rn("newsTitle"))
	getLineCnt = retVal 
End Function  %>
<% Select Case userType
Case "C" %><!--#include file="clientBottom.asp"-->
<% Case "V" %><!--#include file="agentBottom.asp"-->
<% End Select %>
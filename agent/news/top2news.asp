<% addLngPathStr = "news/" %>
<!--#include file="lang/top2news.asp" -->
<% 
Sub doTopNews
  

If userType = "V" Then 
	MainDoc = "ventas"
ElseIf userType = "C" Then
	MainDoc = "clientes"
End If 

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetObjectData" & Session("ID")
cmd.Parameters.Refresh()
cmd("@ObjType") = "S"
cmd("@ObjID") = 1
cmd("@UserType") = userType
set ra = cmd.execute()

set rn = server.CreateObject("ADODB.RecordSet")
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetTopXNews" & Session("ID")
cmd.Parameters.Refresh()
cmd("@LanID") = Session("LanID")
cmd("@TopX") = ra("TopX")
cmd("@MaxShortTextLen") = ra("MaxShortTextLen")
rn.open cmd, , 3, 1

If Not rn.Eof Then

strContent = ra("ObjContent")
If gettop2newsLngStr("DtxtMore") <> "" Then strContent = Replace(strContent, "{txtMore}", gettop2newsLngStr("DtxtMore"))
strContent = Replace(strContent, "{newsTitle}", Server.HTMLEncode(txtNewss))
strContent = Replace(strContent, "{SelDes}", SelDes)
strContent = Replace(strContent, "{rtl}", Session("rtl"))
strContent = Replace(strContent, "{ImgMaxSize}", ra("ImgMaxSize"))
strContent = Replace(strContent, "{dbName}", Session("olkdb"))
If Session("rtl") = "" Then
	strContent = Replace(strContent, "{txtMoreAlign}", "right")
Else
	strContent = Replace(strContent, "{txtMoreAlign}", "left")
End If

strHdr = Left(strContent, InStr(strContent, "<!--startloop-->")-1)
strCnt = Mid(strContent, InStr(strContent, "<!--startloop-->")+16, InStr(strContent, "<!--endloop-->")-InStr(strContent, "<!--startloop-->")-16)
strSep = Mid(strContent, InStr(strContent, "<!--startsep-->")+15, InStr(strContent, "<!--endsep-->")-InStr(strContent, "<!--startsep-->")-15)
strBtm = Right(strContent, Len(strContent)-InStr(strContent, "<!--endsep-->")-13)


strContent = strCnt
%>
<%=strHdr%>
<% do while not rn.eof
If rn("newsImg") <> "" Then
	Pic = rn("newsImg")
Else
	Pic = "n_a.gif"
End If
strCnt = strContent
strCnt = Replace(strCnt, "{newsName}", rn("newsTitle"))
strCnt = Replace(strCnt, "{newsIndex}", rn("newsIndex"))
strCnt = Replace(strCnt, "{newsDate}", FormatDate(rn("newsDate"), True))
strCnt = Replace(strCnt, "{newsShortText}", rn("newsSmallText"))
strCnt = Replace(strCnt, "{newsPic}", Pic)
%>
<%=strCnt%>
<% If CInt(rn("newsID")) < CInt(ra("TopX")) Then %><%=strSep%><% End If %>
<% rn.movenext
loop %>
<%=strBtm%>
<% 
strContent = ""
strHdr = ""
strCnt = ""
strSep = ""
strBtm = ""
set ra = Nothing
End If
set rn = nothing
End Sub %>
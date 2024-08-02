<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="../chkLogin.asp" -->
<!--#include file="../myHTMLEncode.asp" -->
<!--#include file="lang/pollView.asp" -->
<html <% If Session("rtl") <> "" Then %>dir="rtl" <% End If %>>

<head>
<%
Dim errmsg

      pollIndex = Request("pollIndex")
      
      set rd = Server.CreateObject("ADODB.recordset")           
sql = "select pollIndex, pollTitle, pollDate, " & _
	  "case when exists(select 'A' from olkpolldb where pollindex = T0.pollIndex and " & _
	  "CardCode = N'" & Session("UserName") & "') Then 'Y' Else 'N' End PollVote " & _
      "from olkpoll T0 where pollIndex = " & pollIndex
	  set rd = conn.execute(sql) 
	  
      set rs = Server.CreateObject("ADODB.recordset")      	  
	  sql = "select LineText, " & _
	  "(select count('A') from olkpolldb where pollindex = T0.pollIndex and pollSelection = T0.pollLineNum) Votes, " & _
	  "Case (select count('A') from olkpolldb where pollindex = T0.pollIndex and pollSelection = T0.pollLineNum) When 0 Then 0 Else " & _
	  "(select count('A') from olkpolldb where pollindex = T0.pollIndex and pollSelection = T0.pollLineNum) *100/ " & _
	  "(select count('A') from olkpolldb where pollIndex = T0.pollIndex) End VotesPercantage, colorIndex from olkpollLines T0 " & _
	  "where pollIndex = " & pollIndex & _
	  " order by lineOrder asc"
      rs.open sql, conn, 3, 1		
%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=getpollViewLngStr("LttlPollDet")%></title>
<style>
<!--
A:LINK {
	text-decoration : none;
	font-size : 10px;
	font-family: Verdana;
	color : #000000;
}
A:VISITED {
	text-decoration : none;
	font-weight : none;
	color : #000000;
}
A:ACTIVE {
	text-decoration : underline;
	font-weight : none;
	color : #ff0000;
}
A:HOVER {
	text-decoration : underline;
	font-weight : none;
	color : #000000;
}
-->
</style>
<style type="TEXT/CSS">
<!--
body {
	scrollbar-base-color: #336699;
	scrollbar-highlight-color: #E1F1FF;
}
-->
</style>
</head>

<body topmargin="0" leftmargin="0" bgcolor="#F5FBFE" onbeforeunload="opener.clearWin();">

<div align="center">
	<table border="0" cellpadding="0" bordercolor="#111111" width="224" id="table16">
		<tr>
			<td width="100%">
			<table border="1" cellpadding="0" cellspacing="0" width="100%" id="table19" bordercolor="#7CBCFC">
				<tr>
					<td bordercolor="#336699">
					<table border="0" cellpadding="0" cellspacing="1" width="100%" id="table20" bordercolor="#336699" style="border-width: 1px;">
						<tr>
							<td bgcolor="#E1F3FD" style="border-left-style: none; border-left-width: medium; border-right-style: none; border-right-width: medium; border-top-style: none; border-top-width: medium">
							<b><font size="1" face="Verdana">&nbsp;<%=getpollViewLngStr("LtxtYourOp")%></font></b></td>
						</tr>
						<tr>
							<td bordercolor="#7CBCFC" style="border-bottom-style: solid; border-bottom-width: 1px; border-left-style: none; border-left-width: medium; border-right-style: none; border-right-width: medium">
							<p align="center"><b><font size="1" face="Verdana"><%=rd("pollTitle")%>
							</font></b></p>
							</td>
						</tr>
						<tr>
							<td bgcolor="#336699" style="border-left-style: none; border-left-width: medium; border-right-style: none; border-right-width: medium; border-top-style: solid; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
							<p align="center"><b><font size="1" face="Verdana"><%=FormatDate(rd("pollDate"), True)%>
							</font></b></p>
							</td>
						</tr>
						<tr>
							<td bgcolor="#F0F8FF" style="border-top-style: solid; border-top-width: 1px; border-left-style: none; border-left-width: medium; border-right-style: none; border-right-width: medium" bordercolor="#7CBCFC">
							<% do while not rs.eof %>
							<font color="black" size="1" face="Verdana" color="#7CBCFC">
							<b>&nbsp;<%=rs("lineText")%></b></font><br>
							<% If rs("VotesPercantage") > 0 then %>
							<img border="0" src='../poll/colo<%=rs("colorIndex")%>.gif' width='<%=rs("VotesPercantage")%>' height="12"><% end if %><font face="verdana" size="1"> 
							%<%=rs("VotesPercantage")%></font><br>
							<% rs.movenext
                            	loop %></td>
						</tr>
						<tr>
							<td bgcolor="#F0F8FF" style="border-left-style: none; border-left-width: medium; border-right-style: none; border-right-width: medium; border-bottom-style: none; border-bottom-width: medium; font-size: 4px">&nbsp;</td>
						</tr>
					</table>
					</td>
				</tr>
			</table>
			</td>
		</tr>
	</table>
</div>

</body>
<% set rd = nothing
set rs = nothing
conn.close %>

</html>

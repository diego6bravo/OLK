<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="../chkLogin.asp" -->
<html <% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>>
<!--#include file="../myHTMLEncode.asp"-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>New Page 2</title>
<link rel="stylesheet" type="text/css" href="../design/<%=Request("SelDes")%>/style/viewer_css.css">
<% set rs = Server.CreateObject("ADODB.recordset") %>
<script language="javascript">
var EnRClickBlk = <% If myApp.EnBlockRClk Then %>true<% Else %>false<% End If %>;
var RClickBlkMsg = '<%=Replace(Replace(myApp.EnBlkRClkMsg, "'", "\'"), VbCrLf, "\n")%>';
</script>
<script language="javascript" src="../rightclick.js"></script>
</head>

<%
If Request("Item") <> "" Then
	sql = "select AliasID from cufd where tableid = 'OITM' and EditType = 'I'"
	set rs = conn.execute(sql)
	i = 0
	do while not rs.eof
		i = i + 1
		unionImg = unionImg & "union(select " & i & " ID, U_" & rs("AliasID") & " Img from oitm where itemcode = N'" & myQSDecode(Request("Item")) & "' and U_" & rs("AliasID") & " is not null) "
	rs.movenext
	loop 
	rs.close
	sql = _
	"(select 0 ID, IsNull(PicturName,'n_a.gif') Img from oitm where itemcode = N'" & myQSDecode(Request("Item")) & "') " & unionImg
	sql = "select * from (" & sql & ") X0"
ElseIf Request("card") <> "" Then
	sql = "select AliasID from cufd where tableid = 'OCRD' and EditType = 'I'"
	set rs = conn.execute(sql)
	i = 0
	do while not rs.eof
		i = i + 1
		unionImg = unionImg & "union(select " & i & " ID, U_" & rs("AliasID") & " Img from OCRD where CardCode = N'" & myQSDecode(Request("card")) & "' and U_" & rs("AliasID") & " is not null) "
	rs.movenext
	loop 
	rs.close
	sql = _
	"(select 0 ID, IsNull(Picture,'n_a.gif') Img from OCRD where CardCode = N'" & myQSDecode(Request("card")) & "') " & unionImg
	sql = "select * from (" & sql & ") X0"

End If
rs.open sql, conn, 3, 1
%>
<body topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" <% If Not rs.eof Then %>onload="showImg(1);"<% End If %>>
<script language="javascript">
imgCount = <%=rs.recordcount%>;
selImg = -1;
var zoom = '';
function setZoom(val) { zoom = val; }

function showImg(img)
{
	if (selImg != -1)
	{
		document.getElementById('td' + selImg).className='TblBackIntoViw';
		document.getElementById('sel' + selImg).value = 'N';
	}
	document.getElementById('td' + img).className='hlt';
	document.getElementById('sel' + img).value = 'Y';

	parent.parent.frames[0].setImg(document.getElementById('img' + img).value);
	selImg = img;
	stopTimer();
}
function nextPic()
{
	if (selImg < imgCount) showImg(selImg+1);
	else showImg(1);
	setScroll();
}
function prevPic()
{
	if (selImg > 1) showImg(selImg-1);
	else showImg(imgCount);
	setScroll();
}
function setScroll()
{
	document.body.scrollLeft = document.getElementById('td' + selImg).offsetLeft;
}
var imgTimerID = 0;
function startTimer()
{
	imgTimerID  = setTimeout("doTimer()", 3000);
}
function stopTimer()
{
	clearTimeout(imgTimerID);
}
function doTimer()
{
	nextPic();
	imgTimerID = setTimeout("doTimer()", 3000);
}
function chkStop()
{
	if (parent.isInPlay() == 'Y') parent.doPlay(false);
}
</script>
<table border="0" cellpadding="0" cellspacing="0" id="table1" width="<%=81*rs.recordcount%>">
	<tr class="TblBackViw">
		<td valign="top">
		<div align="left">
			<table border="0" cellpadding="0" id="table2">
				<tr>
					<td>
					<table border="0" cellpadding="0" width="100%" id="table3">
						<tr>
							<% do while not rs.eof %>
							<input type="hidden" id="img<%=rs.bookmark%>" value="<%=rs("Img")%>">
							<input type="hidden" id="sel<%=rs.bookmark%>" value="N">
							<td class="TblBackIntoViw" width="70" id="td<%=rs.bookmark%>" onmouseover="this.className = 'hlt';" onmouseout="if(sel<%=rs.bookmark%>.value=='N')this.className = 'TblBackIntoViw';" style="padding: 4px; cursor: hand" onclick="javascript:chkStop();showImg(<%=rs.bookmark%>);">
							<p align="center">
							<img border="1" src="../pic.aspx?filename=<%=rs("Img")%>&dbName=<%=Session("olkdb")%>&MaxSize=60"></td>
							<% rs.movenext
							loop %>
						</tr>
					</table>
					</td>
				</tr>
			</table>
		</div>
		</td>
	</tr>
</table>

</body>

</html>
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
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Cambio de imagen</title>
<% set rs = Server.CreateObject("ADODB.recordset") %>
<script language="javascript">
var EnRClickBlk = <% If myApp.EnBlockRClk Then %>true<% Else %>false<% End If %>;
var RClickBlkMsg = '<%=Replace(Replace(myApp.EnBlkRClkMsg, "'", "\'"), VbCrLf, "\n")%>';
var varImg;
var zoom = '';
function setZoom(val) { zoom = val; }

function setImg(img)
{
	varImg = img;
	resizeImg();
}
function resizeImg()
{
	isZoom = (zoom != null && zoom == '') ? 'N' : 'Y';
	MaxSizeH = document.body.clientHeight-14;
	MaxSizeW = document.body.clientWidth-14;
	if (zoom != '' && zoom != null)
	{
		MaxSizeH = MaxSizeH*(parseInt(zoom)/100)
		MaxSizeW = MaxSizeW*(parseInt(zoom)/100)
	}
	
	var MaxDif;

	if (MaxSizeH < MaxSizeW) 
	{
		MaxDif = MaxSizeW-MaxSizeH;
	}
	else if (MaxSizeH >= MaxSizeW)
	{
		MaxDif = MaxSizeH-MaxSizeW
	}
	mainImg.src='../pic.aspx?filename=' + varImg + '&dbName=<%=Session("olkdb")%>&MaxSize=' + parseInt(MaxSizeH) + '&wide=Y&MaxDif=' + parseInt(MaxDif) + '&isZoom=' + isZoom;
}
</script>
<script language="javascript" src="../rightclick.js"></script>
<link rel="stylesheet" type="text/css" href="../design/<%=Request("SelDes")%>/style/viewer_css.css">
</head>
<body topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" onresize="resizeImg();">
<table border="0" cellpadding="0" width="100%" height="100%" id="table1">
	<tr class="BrdViw">
		<td style="padding: 5px">
		<div align="center">
			<table border="0" cellpadding="0" width="468" cellspacing="1" id="table2">
				<tr>
					<td>
					<table id="table3" height="344" cellSpacing="0" cellPadding="0" width="100%" border="0">
						<tr>
							<td>
							<p align="center">
							<img border="1" name="mainImg" src="../pic.aspx?filename=n_a.gif&dbName=<%=Session("olkdb")%>&MaxSize=100"></td>
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
<% set rs = nothing
conn.close %>

</html>
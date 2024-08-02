<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="../chkLogin.asp" -->
<!--#include file="lang/ImageList.asp" -->
<html <% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title></title>
<!--#include file="../myHTMLEncode.asp"-->
<link rel="stylesheet" type="text/css" href="../design/<%=Request("SelDes")%>/style/viewer_css.css">
</head>
<% set rs = Server.CreateObject("ADODB.recordset") %>
<script language="javascript">
var EnRClickBlk = <% If myApp.EnBlockRClk Then %>true<% Else %>false<% End If %>;
var RClickBlkMsg = '<%=Replace(Replace(myApp.EnBlkRClkMsg, "'", "\'"), VbCrLf, "\n")%>';
function isInPlay() { return isPlay.value; }
function doPlay(play)
{
	if (play)
	{
		isPlay.value = 'Y';
		picLst.startTimer();
		btnPlay.src='../design/<%=Request("SelDes")%>/images/btnPlayDis.jpg';
		btnStop.src='../design/<%=Request("SelDes")%>/images/btnStop.jpg';
	}
	else
	{
		isPlay.value = 'N';
		picLst.stopTimer();
		btnPlay.src='../design/<%=Request("SelDes")%>/images/btnPlay.jpg';
		btnStop.src='../design/<%=Request("SelDes")%>/images/btnStopDis.jpg';
	}
}
function doZoom(val)
{
	picLst.setZoom(val);
	parent.frames[0].setZoom(val);
	parent.frames[0].resizeImg();
}
</script>
<script language="javascript" src="../rightclick.js"></script>
<body topmargin="0" leftmargin="15" rightmargin="0" bottommargin="0">
<input type="hidden" id="isPlay" value="N">
<div align="left">
	<table border="0" cellpadding="0" width="100%" cellspacing="0" id="table1">
		<tr class="BrdViw">
			<td style="padding: 1px">
			<table border="0" cellpadding="0" width="100%" id="table2">
				<tr class="TblBackViw">
					<td>
					<table border="0" cellpadding="0" width="100%" id="table3">
						<tr class="TblBackViw">
							<td style="padding: 1px" valign="top">
							<iframe name="picLst" id="picLst" width="100%" height="104" src="iframe.asp?item=<%=Request("item")%>&card=<%=Request("card")%>&SelDes=<%=Request("SelDes")%>" border="0" frameborder="0">
							Your browser does not support inline frames or is currently configured not to display inline frames.
							</iframe></td>
							<td width="153" style="padding: 1px">
							<div align="center">
								<table border="0" cellpadding="0" width="100%" cellspacing="1" id="table4">
									<tr class="TblBrdBackViw">
										<td>
										<table border="0" cellpadding="0" width="100%" cellspacing="1" id="table5">
											<tr class="TblBackIntoViw">
												<td style="padding: 2px">
												<table border="0" cellpadding="0" width="100%" cellspacing="1" id="table6">
													<tr>
														<td>
															<p align="center">
															<select size="1" name="cmbZoom" onchange="doZoom(this.value);">
															<option>- 
															<%=getImageListLngStr("LtxtAutomatic")%> -
															</option>
															<option value="150">
															150%
															</option>
															<option value="200">
															200%
															</option>
															<option value="400">
															400%
															</option>
															</select></p>
														</td>
													</tr>
													<tr>
														<td align="center">
														<a href="javascript:if(isPlay.value=='Y')doPlay(false);picLst.prevPic();">
														<img border="0" src="../design/<%=Request("SelDes")%>/images/btn<% If Session("rtl") = "" Then %>Ant<% Else %>Sig<% End If %>.jpg" width="30" height="30" alt="<%=getImageListLngStr("LtxtPrev")%>"></a>
														<a href="javascript:if(isPlay.value=='Y')doPlay(false);picLst.nextPic();">
														<img border="0" src="../design/<%=Request("SelDes")%>/images/btn<% If Session("rtl") = "" Then %>Sig<% Else %>Ant<% End If %>.jpg" width="30" height="30" alt="<%=getImageListLngStr("LtxtNext")%>"></a></td>
													</tr>
													<tr>
														<td align="center">
														<img border="0" id="btnPlay" src="../design/<%=Request("SelDes")%>/images/btnPlay.jpg" width="30" height="30" alt="<%=getImageListLngStr("LtxtStart")%>" onmouseover="if(isPlay.value=='N')this.style.cursor='hand';" onmouseout="this.style.cursor='';" onclick="javascript:if(isPlay.value=='N')doPlay(true);">
														<img border="0" id="btnStop" src="../design/<%=Request("SelDes")%>/images/btnStopDis.jpg" width="30" height="30" alt="<%=getImageListLngStr("LtxtStop")%>" onmouseover="if(isPlay.value=='Y')this.style.cursor='hand';" onmouseout="this.style.cursor='';" onclick="javascript:if(isPlay.value=='Y')doPlay(false);"></td>
													</tr>
												</table>
												</td>
											</tr>
										</table>
										</td>
									</tr>
								</table>
							</div>
							</td>
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
<% set rs = nothing
conn.close %>
</html>
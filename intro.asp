<!--#include file="lang.asp"-->
<!--#include file="langIndex.inc"-->
<!--#include file="lang/intro.asp" -->
<html <% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>>

<head>
<meta content="text/html; charset=utf-8" http-equiv="Content-Type">
<title>OnlineKit</title>
<script language="javascript">
var curDir = '<% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>';
var rtl = '<%=Session("rtl")%>';
var myLng = '<% For i = 0 to UBound(myLanIndex) %><% If i > 0 Then %>, <% End If %><%=myLanIndex(i)(0)%>{S}<span lang="<%=myLanIndex(i)(2)%>" dir="<%=myLanIndex(i)(3)%>"><%=myLanIndex(i)(1)%></span><% Next %>';
</script>
<script language="javascript" src="general.js"></script>
<script language="javascript" src="js_lang_sel.js"></script>
<style>
.OlkBtn {
	padding-<% If Session("rtl") = "" Then %>left<% Else %>right<% End If %>: 14px;
	cursor: pointer;
	font-family: Arial, Helvetica, sans-serif;
	font-size: x-small;
	font-weight: bold;
	color: #05649E;
	background-image: url('images/btn.jpg');
	background-repeat: no-repeat;
	height: 18px;
	opacity:0.75;
	filter:alpha(opacity=75);
}
.OlkBtnOver {
	padding-<% If Session("rtl") = "" Then %>left<% Else %>right<% End If %>: 14px;
	cursor: pointer;
	font-family: Arial, Helvetica, sans-serif;
	font-size: x-small;
	font-weight: bold;
	color: #05649E;
	background-image: url('images/btn.jpg');
	background-repeat: no-repeat;
	height: 18px;
}
</style>
</head>

<body topmargin="0" style="padding-top: 100px;">
<%
dim fs
set fs=Server.CreateObject("Scripting.FileSystemObject")

%>
<center>
	<div style="background-image: url('images/intro.jpg'); background-repeat: no-repeat; width: 448px; height: 267px; padding-top: 100px;">
		<div style="height: 75px;">
		<table cellpadding="0" cellspacing="1" style="width: 200px;">
			<% If fs.FolderExists(Server.MapPath("admin/")) Then %>
			<tr>
				<td class="OlkBtn" onmouseover="this.className='OlkBtnOver';" onmouseout="this.className='OlkBtn';" onclick="window.location.href='admin/';"><%=UCase(getintroLngStr("LtxtAdmin"))%></td>
			</tr>
			<% End If %>
			<% If fs.FolderExists(Server.MapPath("agent/")) Then %>
			<tr>
				<td class="OlkBtn" onmouseover="this.className='OlkBtnOver';" onmouseout="this.className='OlkBtn';" onclick="window.location.href='agent/';"><%=UCase(getintroLngStr("DtxtAgent"))%></td>
			</tr>
			<% End If %>
			<% If fs.FolderExists(Server.MapPath("client/")) Then %>
			<tr>
				<td class="OlkBtn" onmouseover="this.className='OlkBtnOver';" onmouseout="this.className='OlkBtn';" onclick="window.location.href='client/';"><%=UCase(getintroLngStr("DtxtClient"))%></td>
			</tr>
			<% End If %>
			<% If fs.FolderExists(Server.MapPath("mobile/")) Then %>
			<tr>
				<td class="OlkBtn" onmouseover="this.className='OlkBtnOver';" onmouseout="this.className='OlkBtn';" onclick="window.location.href='mobile/';"><%=UCase(getintroLngStr("DtxtPocket"))%></td>
			</tr>
			<% End If %>
		</table>
		</div>
		<table cellpadding="0" cellspacing="1" style="width: 200px;">
			<tr>
				<td align="center">
				<table>
					<tr>
						<td><font color="#C0C0C0" face="Verdana" size="1">
						<a href="http://www.topmanage.com.pa/">
						<font color="#C0C0C0">TopManage</font></a> ® 2002 - 2011
						<br><%=getintroLngStr("DtxtEMail")%>: <a href="mailto:info@topmanage.com.pa">
						<font color="#C0C0C0">info@topmanage.com.pa</font></a><font color="#C0C0C0">
						<br><%=getintroLngStr("DtxtPhone")%>: <span dir="ltr">(507) 300-7200</span></font></font></td>
					</tr>
				</table>
				</td>
			</tr>
		</table>
		<div align="right">
	   <table cellpadding="0" cellspacing="0" border="0">
	   	<tr>
			<td width="27" align="right" style="padding-right: 10px; padding-top: 2px;" valign="top">
				<table cellpadding="0" cellspacing="0" border="0" onclick="return !showSelectLang(this, event);" style="border: solid 1px #016AA3;cursor: hand;">
				<tr>
					<td width="16" height="16" align="center">
					<font size="1" face="Verdana" color="#016AA3"><%=LEFT(UCase(Session("myLng")), 2)%></font></td>
				</tr>
				</table>
			</td>
		</tr>
	   </table>
	   </div>

	</div>
</center>
<form action="intro.asp" method="post" name="frmChangeLng">
	<input name="newLng" type="hidden" value="">
</form>
<script> 
var jsLangCol1 = '#016AA3';
var jsLangCol2 = '#FFFFFF';
var jsLangCol3 = '#DEEBF3';
var jsLangCol4 = '#DEEBF3';
var jsLangCol5 = '#787878';
var jsLangRev = false;

if (typeof doNoLang == 'undefined')
{
	doSelLang();
	document.onclick=clearSelectLang;
}
</script>

</body>

</html>
<% set fs=nothing %>
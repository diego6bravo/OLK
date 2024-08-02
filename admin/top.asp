<!--#include file="chkLogin.asp" -->
<!--#include file="lang.asp"-->
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<!--#include file="lang/top.asp" -->
<html <% If Session("rtl") <> "" Then %>dir="rtl"<% End If %>>
<% If userType = "C" Then
	response.redirect "default.asp?cmd=home"
ElseIf userType = "V" Then
	response.redirect "agent.asp"
End If %>
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

set rs = Server.CreateObject("ADODB.RecordSet")
set rm = Server.CreateObject("ADODB.RecordSet")
OLKVerStr = myApp.OLKVersion
R3VerStr = myApp.R3Version

If mySession.IsDatabaseLoaded Then
	CmpName = mySession.GetCompanyName
End If %>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>TopManage OnLineKit - <%=gettopLngStr("LTtlAdmin")%></title>
<!--#include file="adminMenus.asp" -->
<script type="text/javascript" src="general.js"></script>
<script type="text/javascript" src="js_msgbox.js"></script>
<script src="http://code.jquery.com/jquery-latest.js"></script>
<script type="text/javascript">
function goLink(Action) { window.location.href = Action; } 
function toggle(e) {
  for (var i = <% If mySession.IsDatabaseLoaded Then %>1<% Else %>3<% End If %>;i<=<% If mySession.IsDatabaseLoaded Then %>5<% Else %>3<% End If %>;i++)
  {
  	if (e != document.getElementById('HideShow' + i))
  	document.getElementById('HideShow' + i).style.display = "none";
  }
  if (e.style.display == "none") {
     e.style.display = "";
  } else {
     e.style.display = "none";
  }
}

var myLng = '<% For i = 0 to UBound(myLanIndex) %><% If i > 0 Then %>, <% End If %><%=myLanIndex(i)(0)%>{S}<span lang="<%=myLanIndex(i)(2)%>" dir="<%=myLanIndex(i)(3)%>"><%=myLanIndex(i)(1)%></span><% Next %>';
</script>
<link rel="stylesheet" type="text/css" href="style/style_admin_<%=Session("style")%>.css">
<link rel="stylesheet" type="text/css" media="all" href="style_cal.css" >
<link type="text/css" href="style/jquery-ui-1.8.14.custom.css" rel="stylesheet" >	
<script type="text/javascript" src="jQuery/js/jquery-1.6.1.min.js"></script>
<% If curPage = "adminRepEdit.asp" Then %>
<link rel="stylesheet" href="js_color_picker_v2.css" media="screen">
<script type="text/javascript" src="color_functions.js"></script>		
<script type="text/javascript" src="js_color_picker_v2.js"></script>
<% End If %>
<script type="text/javascript">
var rtl = '<%=Session("rtl")%>';
</script>
<script type="text/javascript" src="js_lang_sel.js"></script>
<script type="text/javascript" src="admin.js"></script>
<script type="text/javascript">
var errEMailValDomain = '<%=gettopLngStr("LerrEMailValDomain")%>';
var errEMailValURL = '<%=gettopLngStr("LerrEMailValURL")%>';
</script>
</head>
<% r = 0
Select Case curPage
	Case "adminSecEdit.asp"
		If Request("UType") = "C" Then onLoad = onLoad & "loadSmallText();"
		onLoad = onLoad & "if (isManual == 'Y') doLoadPreview();"
	Case "adminDefObjEdit.asp"
		onLoad = onLoad & "doLoadPreview();"
End Select  %>
<body onfocus="javascript:chkWin();" onload="<%=onload%>" onscroll="setMsgBoxPos();" onresize="setMsgBoxPos();" style="margin-top: 0px; margin-left: 0px; margin-right: 0px; background-color: #C0EEFE;">

<table border="0" cellpadding="0" cellspacing="0" width="100%">
  	<tr>
		<td width="9%">
		<img src="images/spacer.gif" width="172" height="1" alt=""/></td>
		<td width="12%">
		<img src="images/spacer.gif" width="232" height="1" alt=""/></td>
		<td width="42%">
		<img src="images/spacer.gif" width="50" height="1" alt=""/></td>
		<td width="35%">
		<img src="images/spacer.gif" width="246" height="1" alt=""/></td>
	</tr>
	<tr>
		<td colspan="2" width="100%">
		<img name="admin_olk_new_r1_c1" src="images/<%=Session("rtl")%>admin_olk_new_r1_c1.jpg" width="404" height="18" border="0" alt=""></td>
		<td rowspan="2" width="100%" background="images/backtop.jpg" <% If Session("rtl") = "rtl/" Then %>align="right"<% End If %>>
		<img name="admin_olk_new_r1_c3" src="images/<%=Session("rtl")%>admin_olk_new_r1_c3.jpg" width="50" height="145" border="0" alt=""></td>
		<td width="35%" align="<% If Session("rtl") <> "rtl/" Then %>right<% Else %>left<% End If %>" background="images/admin_olk_new_r1_c4_bg.jpg">
		<font face="Verdana" size="2" color="#1A759E"><i><b><%=gettopLngStr("LTtlAdmin")%></b></i></font></td>
	</tr>
	<tr>
		<td colspan="2" width="100%" background="images/<%=Session("rtl")%>admin_olk_new_r2_c1.jpg">
		<table border="0" cellpadding="0" width="100%" id="table2">
			<tr>
				<td><br>
				<br>
				&nbsp;</td>
			</tr>
			<tr>
				<td><img border="0" src="images/point.gif" width="1" height="2"></td>
			</tr>
			<tr>
				<td>
				<p<% If Session("rtl") = "rtl/" Then %> align="right"<% End If %>><font face="Verdana" size="1" color="#2C96D4"><% If CmpName <> "" Then %><span dir="ltr">[<%=Session("olkdb")%>]&nbsp;<%=CmpName%></span><% End If %>&nbsp;</font></td>
			</tr>
		</table>
		</td>
		<td width="246" height="127" background="images/<%=Session("rtl")%>img_top.jpg" align="right" valign="bottom">
		&nbsp;
		</td>
	</tr>
	<tr>
		<td background="images/<%=Session("rtl")%>back_iz.jpg" width="9%" valign="top">
<table border="0" cellpadding="0" width="98%" id="table1" cellspacing="1">
	<tr onclick="javascript:goLink('admin.asp');" style="cursor: hand">
		<td bgcolor="#D9F5FF">
		<img border="0" src="images/inicio_icon.gif" width="15" height="17"><font color="#31659C"> </font>
		<b><font style="font-size: 9pt" face="Tahoma" color="#31659C">
		<%=gettopLngStr("LtxtHome")%></font></b></td>
	</tr>
	<tr>
		<td>
		<img border="0" src="images/trans.gif" width="3" height="1"></td>
	</tr>
	<% If mySession.IsDatabaseLoaded Then %>
	<tr>
		<td bgcolor="#D9F5FF"><div style="cursor: hand" onclick="toggle(HideShow1);">
			<img border="0" src="images/generales_icon.gif" width="15" height="17"><b><font face="Tahoma" color="#31659C" style="font-size: 9pt"> 
			<%=gettopLngStr("LtxtGeneral")%></font></b></div></td>
	</tr>
	<tr>
		<td <% If Not ShowGeneral Then %>style="display: none"<% End If %> id=HideShow1 bgcolor="#BFEEFE">
		<% DisplayMenuItems(mnuGen) %>
		</td>
	</tr>
	<tr>
		<td bgcolor="#D9F5FF"><b>
		<font style="font-size: 9pt" face="Tahoma" color="#31659C"><div style="cursor: hand" onclick="toggle(HideShow2);">
			<img border="0" src="images/ambiente_icon.gif" width="15" height="17"> 
			<%=gettopLngStr("LtxtAmbient")%></div></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#BFEEFE" <% If Not ShowEnvir Then %>style="display: none"<% End If %> id=HideShow2>
				<table border="0" cellpadding="0" class="TablaMenuSection" onmouseover="bgColor='#D9F5FF'; this.style.cursor='hand';; r8.style.backgroundColor='#2C96D4'" onmouseout="bgColor='#BFEEFE';; r8.style.backgroundColor='#E5F7FB'"  width="100%" id="table9">
				<% DisplayMenuItems(mnuEnvir) %>
				</td>
	</tr>
	<% End If %>
	<tr>
		<td bgcolor="#D9F5FF"><b>
		<font style="font-size: 9pt" face="Tahoma" color="#31659C"><div style="cursor: hand" onclick="toggle(HideShow3);">
			<img border="0" src="images/acceso_icon.gif" width="15" height="17"> 
			<%=gettopLngStr("LtxtAccess")%></div></font></b></td>
	</tr>
	<tr >
		<td bgcolor="#BFEEFE" <% If Not ShowAcc Then %>style="display: none"<% End If %> id=HideShow3>
				<% DisplayMenuItems(mnuAcc) %>
				</td>
	</tr>
	<% If mySession.IsDatabaseLoaded Then %>
	<tr>
		<td bgcolor="#D9F5FF"><b>
		<font style="font-size: 9pt" face="Tahoma" color="#31659C">
		<div style="cursor: hand" onclick="toggle(HideShow4);">
		<img border="0" src="images/personalizacion_icon.gif" width="15" height="17"> 
		<%=gettopLngStr("LtxtCustomize")%></div></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#BFEEFE" <% If Not ShowPer Then %>style="display: none"<% End If %> id=HideShow4>
				<% DisplayMenuItems(mnuPers) %>
				</td>
	</tr>
	<tr>
		<td bgcolor="#D9F5FF"><b>
		<font style="font-size: 9pt" face="Tahoma" color="#31659C">
		<div style="cursor: hand" onclick="toggle(HideShow5);">
			<img border="0" src="images/reporte_icon.gif" width="15" height="17"> 
			<%=gettopLngStr("LtxtReports")%></div></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#BFEEFE" <% If Not ShowRep Then %>style="display: none"<% End If %> id=HideShow5>
				<% DisplayMenuItems(mnuRep) %>
				</td>
	</tr>
	<tr>
		<td>
		<img border="0" src="images/trans.gif" width="3" height="1"></td>
	</tr>
	<% End If %>
	<tr onclick="javascript:goLink('adminSystem.asp');" style="cursor: hand">
		<td bgcolor="#D9F5FF">
		<img border="0" src="images/system_icon.gif" width="15" height="17"><b><font style="font-size: 9pt" face="Tahoma" color="#31659C"> 
		<%=gettopLngStr("LtxtSystem")%></font></b></td>
	</tr>
	<tr>
		<td>
		<img border="0" src="images/trans.gif" width="3" height="1"></td>
	</tr>
	<tr onclick="javascript:goLink('adminUpdate.asp');" style="cursor: hand">
		<td bgcolor="#D9F5FF">
		<img border="0" src="images/actualizar_icon.gif" width="15" height="17"><b><font style="font-size: 9pt" face="Tahoma" color="#31659C"> 
		<%=gettopLngStr("LtxtUpdates")%></font></b></td>
	</tr>
	<tr>
		<td>
		<img border="0" src="images/trans.gif" width="3" height="1"></td>
	</tr>
	<tr onclick="javascript:goLink('adminLicInf.asp');" style="cursor: hand">
		<td bgcolor="#D9F5FF">
		<img border="0" src="images/licencia_icon.gif" width="15" height="17"><b><font style="font-size: 9pt" face="Tahoma" color="#31659C"> 
		<%=gettopLngStr("LtxtLicence")%></font></b></td>
	</tr>
	<tr>
		<td>
		<img border="0" src="images/trans.gif" width="3" height="1"></td>
	</tr>
	<tr onclick="javascript:goLink('default.asp?logout=Y');" style="cursor: hand">
		<td bgcolor="#D9F5FF">
		<img border="0" src="images/sesion_icon.gif" width="15" height="17"><b><font style="font-size: 9pt" face="Tahoma" color="#31659C"> 
		<%=gettopLngStr("LtxtSignOut")%></font></b></td>
	</tr>
	</table>
		<p>&nbsp;</p>
		<p>&nbsp;</p>
		<p>&nbsp;</p>
		<p>&nbsp;</p>
		<p>&nbsp;</p>
		<p>&nbsp;</p>
		<p>&nbsp;</p>
		<p>&nbsp;</p>
		<p>&nbsp;</td>
		<td colspan="3" style="background-image: Url('images/<%=Session("rtl")%>back.gif'); <% If Session("rtl") <> "" Then %>background-position: top right;<% End If %>" valign="top">
		<table border="0" cellpadding="0" width="100%" id="table1">
			<tr>
				<td><font face="Verdana" size="3"><b>&nbsp;<font color="#2B96D3"> <% If Session("rtl") = "" Then %>//<% Else %>\\<% End If %> 
				<%=CurSection%></font></b></font></td>
				<td align="<% If Session("rtl") <> "rtl/" Then %>right<% Else %>left<% End If %>">
				<table cellpadding="0" cellspacing="0" border="0" bgcolor="#4693B9" style="cursor: hand" onclick="return !showSelectLang(this, event);">
					<tr>
						<td width="16" height="16" align="center">
						<font size="1" face="Verdana" color="#FFFFFF"><%=LEFT(UCase(Session("myLng")), 2)%></font></td>
					</tr>
				</table>
				</td>
				<td width="4"><font size="1" face="Verdana">&nbsp;</font></td>
			</tr>
			<tr>
				<td colspan="3">
				</td>
			</tr>
		</table>
<% 

Sub DisplayMenuItems(ByVal mnuArr)
ArrVal = Split(mnuArr, ", ")
For i = 0 to UBound(ArrVal)
mnuItm = Split(ArrVal(i),"|") %>
<% If curPage <> mnuItm(0) or mnuShowNormal Then %>
<table border="0" cellpadding="0" class="TablaMenuSection" onmouseover="bgColor='#D9F5FF'; this.style.cursor='hand';; r<%=r%>.style.backgroundColor='#2C96D4'" onmouseout="bgColor='#BFEEFE';; r<%=r%>.style.backgroundColor='#E5F7FB'"  width="100%">
<tr onclick="javascript:goLink('<%=mnuItm(0)%>');">
	<td class="TablaMenusmall" width="4" id="r<%=r%>" >&nbsp;</td>
	<td><%=mnuItm(1)%></td>
</tr>
</table>
<% r = r + 1
Else %>
<table border="0" cellpadding="0" class="TablaMenuSection" bgcolor="#D9F5FF" width="100%">
<tr>
	<td width="4" class="TablaMenusmall" style="background-color: #2C96D4; ">&nbsp;</td>
	<td><%=mnuItm(1)%></td>
</tr>
</table>
<% End If %><% 
Next
End Sub %>
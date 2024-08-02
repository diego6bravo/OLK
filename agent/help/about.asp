<!--#include file="lang/about.asp" -->
<!--#include file="../myHTMLEncode.asp"-->
<html <% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>TopManage OLK v<%=Request("Version")%></title>
</head>

<body topmargin="0" leftmargin="0" bgcolor="#EFF8FF">
<table border="0" cellpadding="0" cellspacing="0" width="520">
<!-- fwtable fwsrc="about_olk.png" fwpage="Page 1" fwbase="about_en.jpg" fwstyle="FrontPage" fwdocid = "1910412634" fwnested="0" -->
  <tr>
   <td><img src="images/spacer.gif" width="1" height="1" border="0" alt=""></td>
   <td><img src="images/spacer.gif" width="518" height="1" border="0" alt=""></td>
   <td><img src="images/spacer.gif" width="1" height="1" border="0" alt=""></td>
   <td><img src="images/spacer.gif" width="1" height="1" border="0" alt=""></td>
  </tr>

  <tr>
   <td colspan="3"><img name="about_en_r1_c1" src="images/about_en_r1_c1.jpg" width="520" height="83" border="0" alt=""></td>
   <td><img src="images/spacer.gif" width="1" height="83" border="0" alt=""></td>
  </tr>
  <tr>
   <td><img name="about_en_r2_c1" src="images/about_en_r2_c1.jpg" width="1" height="225" border="0" alt=""></td>
   <td background="images/about_en_r2_c2.jpg" width="518" height="225" valign="top">
   <table border="0" cellpadding="0" width="100%" id="table1">
			<tr>
				<td>
				<p align="center">
    <b><font size="1" face="Verdana"><%=getaboutLngStr("LttlAbout")%><%=Request("Version")%></font></b></td>
			</tr>
			<tr>
				<td>
				<div align="center">
					<table border="0" cellpadding="0" width="446" cellspacing="0" id="table2">
						<tr>
							<td>
							<p align="justify"><font face="Verdana" size="1">
							<%=myHTMLDecode(getaboutLngStr("LttlNote"))%></font></td>
						</tr>
					</table>
				</div>
				</td>
			</tr>
		</table>
   </td>
   <td rowspan="2"><img name="about_en_r2_c3" src="images/about_en_r2_c3.jpg" width="1" height="254" border="0" alt=""></td>
   <td><img src="images/spacer.gif" width="1" height="225" border="0" alt=""></td>
  </tr>
  <tr>
   <td colspan="2" width="519" height="29" background="images/<%=Session("rtl")%>about_en_r3_c1.jpg" align="center">
   <table cellpadding="0" cellspacing="0" border="0" width="100%" style="padding-left: 10px; padding-right: 10px">
		<tr>
			<td>
			<font face="Verdana" size="1"><%=myHTMLDecode(getaboutLngStr("LtxtFooter"))%>&nbsp;<b><a target="_blank" href="http://www.topmanage.com.pa"><font color="#000000">
    		http://www.topmanage.com.pa</font></a></b></font>
			</td>
			<td style="font-size: xx-small; padding-left: 10px; padding-right: 10px">
			<input type="button" name="btnClose" value="<%=getaboutLngStr("DtxtAccept")%>" style="font-size: xx-small; color: #01669A;background-color:#F0F6FE; border-width: 1px; border-color: #01669A; width: 100px; height: 20px;" onclick="javascript:window.close();"></td>
		</tr>
	</table>
   </td>
   <td><img src="images/spacer.gif" width="1" height="29" border="0" alt=""></td>
  </tr>
  <tr>
   <td colspan="3"><img name="about_en_r4_c1" src="images/about_en_r4_c1.jpg" width="520" height="10" border="0" alt=""></td>
   <td><img src="images/spacer.gif" width="1" height="10" border="0" alt=""></td>
  </tr>
</table>

</body>

</html>
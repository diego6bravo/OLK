<!--#include file="top.asp" -->
<!--#include file="lang/adminNews.asp" -->
<!--#include file="adminTradSubmit.asp"-->

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style type="text/css">
.style1 {
	border-bottom-style: none;
	border-bottom-width: medium;
}
.style2 {
	border-left-width: 1px;
	border-right-width: 1px;
	border-top: medium none #C0C0C0;
	border-bottom-width: 1px;
}
.style3 {
	text-align: center;
	border-left-width: 1px;
	border-right-width: 1px;
	border-top: medium none #C0C0C0;
	border-bottom-width: 1px;
}
.style4 {
	border-bottom-style: none;
	border-bottom-width: medium;
	text-align: center;
	color: #31659C;
}
.style5 {
	text-align: center;
	color: #31659C;
}
.style6 {
	border-bottom-style: none;
	border-bottom-width: medium;
	font-weight: bold;
	text-align: center;
}
.style7 {
	text-align: center;
	font-weight: normal;
	color: #31659C;
}
.style8 {
	border-bottom-style: none;
	border-bottom-width: medium;
	text-align: center;
	font-size: xx-small;
	color: #31659C;
}
</style>
</head>

<script language="javascript">
function Start(page, w, h, s) {
OpenWin = this.open(page, "CtrlWindow", "toolbar=no,menubar=no,location=no,scrollbars="+s+",resizable=no, width="+w+",height="+h);
}
</script>
<% 
conn.execute("use [" & Session("OLKDB") & "]")
sql = "select newsIndex, newsDate, newsTitle, Status from olknews where Status <> 'D' order by newsdate desc"
set rs = conn.execute(sql)
%>
<table border="0" cellpadding="0" width="100%" id="table3">
	<tr>
		<td bgcolor="#E1F3FD"><b><font color="#FFFFFF">&nbsp;</font><font color="#31659C" face="Verdana" size="1"><%=getadminNewsLngStr("LttlNews")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1"> </font>
		<font face="Verdana" size="1" color="#4783C5"><%=getadminNewsLngStr("LttlNewsNote")%></font></td>
	</tr>
	<form method="post" action="adminSubmit.asp" name="frmNews">
	<tr>
		<td>
		<table border="0" cellpadding="0" id="table6" style="font-family: Verdana; font-size: 10px; width: 100%;">
			<tr>
				<td align="center" bgcolor="#E2F3FC" class="style1" style="width: 16px">
				&nbsp;</td>
				<td bgcolor="#E2F3FC" class="style8" style="width: 100px">
				<p class="style5">
				<font face="Verdana" size="1"><strong><%=getadminNewsLngStr("DtxtDateOfPub")%></strong></font></td>
				<td bgcolor="#E2F3FC" class="style6">
				<p class="style7">
				<font size="1"><strong><%=getadminNewsLngStr("DtxtTitle")%></strong></font></td>
				<td bgcolor="#E2F3FC" class="style4" style="width: 100px">
				<strong><font size="1"><%=getadminNewsLngStr("DtxtActive")%></font></strong></td>
				<td align="center" bgcolor="#E2F3FC" class="style1" style="width: 16px">
				&nbsp;</td>
			</tr>
			<% do while not rs.eof %>
			<input type="hidden" name="newsIndex" value="<%=rs("newsIndex")%>">
			<tr>
				<td class="style2" bgcolor="#F3FBFE" style="width: 16px">
				<a href="adminNewsEdit.asp?newsIndex=<%=rs("newsIndex")%>">
				<img border="0" src="images/<%=Session("rtl")%>flechaselec.gif" width="15" height="13"></a></td>
				<td class="style2" bgcolor="#F3FBFE" style="width: 100px"><font color="#4783C5" size="1"><%=FormatDate(rs("newsDate"), True)%>&nbsp;</font></td>
				<td class="style2" bgcolor="#F3FBFE">
				<table cellpadding="0" cellspacing="0" border="0" width="100%">
					<tr>
						<td valign="middle"><font color="#4783C5" size="1"><%=rs("newsTitle")%></font>
						</td>
						<td width="16">
						<a href="javascript:doFldTrad('News', 'newsIndex', <%=rs("newsIndex")%>, 'alterNewsTitle', 'T', null);"><img src="images/trad.gif" alt="<%=getadminNewsLngStr("DtxtTranslate")%>" border="0"></a></td>
					</tr>
				</table>
				</td>
				<td class="style3" bgcolor="#F3FBFE" style="width: 100px" align="center">
				<input type="checkbox" class="noborder" name="Status<%=rs("newsIndex")%>" value="A" <% If rs("Status") = "A" Then %>checked<% End If %>></td>
				<td class="style2" bgcolor="#F3FBFE" style="width: 16px">
				<a href="javascript:if(confirm('<%=getadminNewsLngStr("LtxtConfDelNews")%>'.replace('{0}', '<%=Replace(rs("newsTitle"), "'", "\'")%>')))window.location.href='adminSubmit.asp?submitCmd=delNews&newsIndex=<%=rs("newsIndex")%>'">
				<img border="0" src="images/remove.gif" width="16" height="16"></a></td>
			</tr>
			<% rs.movenext
			loop %>
		</table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table5">
			<tr>
				<td width="77">
				<input type="submit" value="<%=getadminNewsLngStr("DtxtSave")%>" name="btnSave" class="OlkBtn"></td>
				<td width="77">
				<input type="button" value="<%=getadminNewsLngStr("DtxtNew")%>" name="B1" class="OlkBtn" onclick="javascript:window.location.href='adminNewsEdit.asp'"></td>
				<td><hr color="#0D85C6" size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
	<input type="hidden" name="submitCmd" value="admNews">
	</form>
</table>
<!--#include file="bottom.asp" -->
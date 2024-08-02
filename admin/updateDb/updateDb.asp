<!--#include file="../myHTMLEncode.asp"-->
<!--#include file="lang/updateDb.asp" -->
<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=getupdateDbLngStr("LttlUpdDB")%></title>
</head>

<body style="text-align: center">
<table border="0" cellpadding="0" cellspacing="0" width="497">
  	<tr>
		<td>
		<p align="center">
		<img src="images/spacer.gif" width="1" height="1" border="0" alt=""></td>
		<td>
		<p align="center">
		<img src="images/spacer.gif" width="495" height="1" border="0" alt=""></td>
		<td>
		<p align="center">
		<img src="images/spacer.gif" width="1" height="1" border="0" alt=""></td>
		<td>
		<p align="center">
		<img src="images/spacer.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr>
		<td colspan="3" height="16" bgcolor="#D7EFFD">
		<p align="left">
		<b><font color="#4783C5" face="Verdana" size="1">&nbsp;<%=getupdateDbLngStr("LttlUpdDB")%></font></b></td>
		<td height="16" bgcolor="#D7EFFD">
		<p align="center">
		<img src="images/spacer.gif" width="1" height="15" border="0" alt=""></td>
	</tr>
	<tr>
		<td background="images/ventana_r2_c1.gif">
		<p align="center">
		<img name="ventana_r2_c1" src="images/ventana_r2_c1.gif" width="1" height="263" border="0" alt=""></td>
		<td bgcolor="#FFFFFF" background="images/ventana_r2_c2.gif">
		<p align="center"><font face="Verdana" size="2">
		<img border="0" src="../images/cargando_gif.gif"></font></p>		
		</td>
		<td background="images/ventana_r2_c3.gif">
		<p align="center">
		<img name="ventana_r2_c3" src="images/ventana_r2_c3.gif" width="1" height="263" border="0" alt=""></td>
		<td>
		<p align="center">
		<img src="images/spacer.gif" width="1" height="263" border="0" alt=""></td>
	</tr>
</table>

<script language="javascript">
window.location.href='update.aspx?dbName=<%=Request("dbName")%>&dbID=<%=Request("dbID")%>'
</script>
</body>

</html>
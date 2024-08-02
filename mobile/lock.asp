<!--#include file="lang.asp"-->
<!--#include file="lang/lock.asp" -->
<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=getlockLngStr("LttlExpSes")%></title>
</head>

<body<% If Session("rtl") <> "" Then %> dir="rtl"<% End If %>>

<div align="center">
  <center>
  <table border="0" cellpadding="0"  bordercolor="#111111" width="349" id="AutoNumber1">
    <tr>
      <td width="345" bgcolor="#F0F8FF">
      <p align="center"><img border="0" src="images/lock.gif"></td>
    </tr>
    <tr>
      <td width="345" bgcolor="#F0F8FF">
      <div align="center">
        <center>
        <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="271" id="AutoNumber2">
          <tr>
            <td width="271">
            <p align="justify"><font face="Verdana" size="1"><%=getlockLngStr("LtxtDescExp")%></font></td>
          </tr>
        </table>
        </center>
      </div>
      </td>
    </tr>
    <tr>
      <td width="345" bgcolor="#F0F8FF">&nbsp;</td>
    </tr>
    <tr>
      <td width="345" bgcolor="#F0F8FF">
      <p align="center"><b><font size="2" face="Verdana">
      <a href="default.asp"><font color="#000000">
      <%=getlockLngStr("LtxtReenter")%></font></a></font></b></td>
    </tr>
    </table>
  </center>
</div>

</body>

</html>
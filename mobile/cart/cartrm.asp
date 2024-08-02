<%@ Language=VBScript %>
<!--#include file="lang/cartrm.asp" -->
<!--#include file="../myHTMLEncode.asp"-->
<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

If Request.Form("Confirm") = "Y" Then
sql = "delete doc1 where lognum = " & Session("RetVal") & " and LineNum = " & Request.Form("Line") & _
" delete [" & Session("OlkDB") & "]..olkSalesLines where LogNum = " & Session("RetVal") & " and LineNum = " & Request("Line")
conn.execute(sql)
conn.close 
response.redirect "../operaciones.asp?cmd=cart"
Else
%><html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="viewport" content="width=280">
<title>Mobile OLK</title>
<script language="javascript" src="general.js"></script>
</head>

<body topmargin="0">

<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
        <tr>
      <td bgcolor="#9BC4FF">
      <table border="0" cellpadding="0" bordercolor="#111111" width="100%" id="AutoNumber1">
        <tr>
          <td width="100%">
          <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><%=getcartrmLngStr("LtxtRemoveItem")%>
          </font></b></td>
        </tr>
        <tr>
          <td width="100%" bgcolor="#75ACFF">
          <p align="center"><b>
          <font size="1" face="Verdana"><%=Replace(getcartrmLngStr("LtxtConfRemItm"), "{0}", Request.QueryString("ICode"))%></font></b></td>
        </tr>
        <tr>
          <td width="100%" style="font-size: 10px">
          &nbsp;</td>
        </tr>
        <tr>
          <td width="100%">
          <div align="center">
            <center>
            <table border="0" cellpadding="0" cellspacing="0"  bordercolor="#111111" width="89">
                <form method="POST" action="cartrm.asp">
              <tr>
                <td>
                  <p align="center">
                  <input border="0" src="../images/ok_icon.gif" name="I1" type="image"></p>
                  <input type="hidden" name="Confirm" value="Y">
                  <input type="hidden" name="Line" value="<%=Request.QueryString("Line")%>">
                </td>
                <td>
                  <p align="center">
                  <img src="../images/spacer.gif"><br><img src="../images/spacer.gif"><br><img src="../images/spacer.gif"><br>
                  <a href="../operaciones.asp?cmd=cart"><img src="../images/x_icon.gif" border="0"></a>
				  </p>
                </td>
              </tr>
                </form>
            </table>
            </center>
          </div>
          </td>
        </tr>
        <tr>
          <td width="100%">
          <p align="center">&nbsp;</td>
        </tr>
      </table>
      </td>
    </tr>
    </table>
  </center>
</div>

<% end if %>
</body>
</html>
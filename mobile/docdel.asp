<!--#include file="lang/docdel.asp" -->
<!--#include file="clearItem.asp"-->
<% if Request("RetVal") <> "" Then Session("RetVal") = Request("RetVal")
If Request("Confirm") = "Y" Then
 sql = "select object from r3_obscommon..tlog where LogNum = " & Session("RetVal")
           set rs = conn.execute(sql)
           sql = "update r3_obscommon..tlog set status = 'B' where lognum = " & Session("RetVal")
           conn.execute(sql)
           Session("RetVal") = "" %>
<table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
    <!-- fwtable fwsrc="Z:\topmanage\logos\originales\pocket_art.png" fwbase="pocket_artpieza1.gif" fwstyle="FrontPage" fwdocid = "742308039" fwnested=""0" -->
    <tr>
      <td bgcolor="#9BC4FF">
      <table border="0" cellpadding="0" bordercolor="#111111" width="100%" id="AutoNumber1">
        <tr>
          <td width="100%">
          <p align="center"><img border="0" src="cart/olklogo_img.gif"></td>
        </tr>
        <tr>
          <td width="100%">
          <p align="center"><b><font size="1" face="Verdana"><%
          Select Case rs("Object")
          	Case "17"
          		Response.Write Replace(getdocdelLngStr("LtxtConfCancel"), "{0}", txtOrdr)
          	Case "23"
          		Response.Write Replace(getdocdelLngStr("LtxtConfCancel"), "{0}", txtQuote)
          End Select %></font></b></td>
        </tr>
        <tr>
          <td width="100%">
          <p align="center">
          <% 
          Select Case Request("redir")
          	Case "pendientes"
          		redir = "cmd=pendientes"
          	Case ""
          		redir = "cmd=datos&card=" & CleanItem(Request("c1"))
          End Select %>
      <a href="operaciones.asp?<%=redir%>">
      <img border="0" src="images/ok_icon.gif"></a></td>
        </tr>
      </table>
      </td>
    </tr>
    </table>
<% Else %><table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
    <!-- fwtable fwsrc="Z:\topmanage\logos\originales\pocket_art.png" fwbase="pocket_artpieza1.gif" fwstyle="FrontPage" fwdocid = "742308039" fwnested=""0" -->
    <tr>
      <td>
      <img src="images/spacer.gif" width="100%" height="1" border="0" alt></td>
    </tr>
    <tr>
      <td bgcolor="#9BC4FF">
      <table border="0" cellpadding="0" bordercolor="#111111" width="100%" id="AutoNumber1">
        <tr>
          <td width="100%">
          <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">&nbsp;</td>
        </tr>
        <tr>
          <td width="100%">
          <p align="center"><b><font face="Verdana" size="1"><%=Replace(getdocdelLngStr("LtxtConfDelDoc"), "{0}", Session("RetVal"))%></font></b></td>
        </tr>
        <tr>
          <td width="100%">
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber2" height="15">
  <tr>
    <td width="100%" height="15" colspan="3" style="font-size: 10px">
    </td>
  </tr>
  <tr>
    <td width="40%" height="15">
      <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
      <a href="operaciones.asp?cmd=cartcancel&redir=<%=Request("redir")%>&confirm=Y&c1=<%=CleanItem(Request("c1"))%>">
      <img border="0" src="images/ok_icon.gif"></a></p>
    </td>
    <td width="20%" height="15">
    &nbsp;</td>
    <td width="40%" height="15">
      <p align="left"><a href="javascript:history.go(-1);"><img border="0" src="images/x_icon.gif"></a></p>
    </td>
  </tr>
</table>
          </td>
        </tr>
      </table>
      </td>
    </tr>
    </table>
<% end if %>
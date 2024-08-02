<!--#include file="lang/about.asp" -->
<div align="center">
  <center>
      <table border="0" cellpadding="0"  bordercolor="#111111" width="100%">
        <tr>
          <td width="100%">
          <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><%=getaboutLngStr("LtxtAbout")%>
          </font></b></td>
        </tr>
        <tr>
          <td width="100%">
          <table border="0" cellpadding="0"  bordercolor="#111111" width="100%">
            <tr>
              <td width="100%">
              <p align="center"><img border="0" src="images/topmanagelogo.gif"></td>
            </tr>
            <tr>
              <td width="100%">
              <p align="center"><font face="Arial" size="2">&nbsp;</font><font color="#000080"><b><font face="Arial Narrow" size="2"><%=getaboutLngStr("LtxtLetsTalkBusiness")%></font></b></font></td>
            </tr>
            <tr>
              <td width="100%">
              <p align="center"><b>
              <font face="Verdana" color="#000080" size="1">Mobile OLK.</font><font face="Verdana" size="2" color="#000080">
              </font><font face="Verdana" color="#000080" size="1">V.&nbsp;<%=myApp.OLKVersion%></font></b></td>
            </tr>
            <tr>
              <td width="100%">
              <p align="center"><b>
              <font face="Verdana" color="#000080" size="1"><%=mySession.GetCompanyName%></font></b></td>
            </tr>
            <tr>
              <td width="100%" style="font-size: 10px">&nbsp;</td>
            </tr>
            <tr>
              <td width="100%">
              <p align="center"><b><font size="1" face="Verdana"><%=getaboutLngStr("DtxtPhone")%>: <span lang="es-pa">
              (507) 300.7200<br>
              Fax: (507) 300.7205</span></font></b></td>
            </tr>
            <tr>
              <td width="100%">
              <p align="center"><b><font face="Arial" size="1">
            TopManage</font></b></td>
            </tr>
            <tr>
              <td width="100%">
              <p align="center"><font face="Arial" size="1">
            <span lang="es-pa"><%=getaboutLngStr("LtxtRights")%> 2002-2012</span></font></td>
            </tr>
          </table>
          </td>
        </tr>
        </table>
  </center>
</div>

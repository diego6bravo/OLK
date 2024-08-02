<!--#include file="lang/searchclients.asp" -->
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
    <tr>
      <td bgcolor="#9BC4FF">
      <table border="0" cellpadding="0" bordercolor="#111111" width="100%" id="AutoNumber1">
        <tr>
          <td width="100%">
          <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><%=getsearchclientsLngStr("LtxtClientSearch")%> 
          </font></b></td>
        </tr>
        <tr>
          <td width="100%">
          <form method="POST" action="operaciones.asp?cmd=searchresult" name="search1">
            <table border="0" cellpadding="0"  bordercolor="#111111" width="100%" id="AutoNumber2">
              <tr>
                <td width="100%" style="font-size: 10px">&nbsp;</td>
              </tr>
              <tr>
                <td width="100%">
                <p align="center"><input type="text" name="string" size="27"></td>
              </tr>
              <tr>
                <td width="100%">
                <div align="center">
                  <center>
                  <table border="0" cellpadding="0" cellspacing="1"  bordercolor="#111111" width="100" id="AutoNumber3">
                    <tr>
                      <td width="63">
                      <input type="submit" value="<%=getsearchclientsLngStr("DtxtSearch")%>" name="B1"></td>
                      <td width="34">
						<% 
						sql = 	"select T0.ID  " & _
								"from OLKCustomSearch T0 " & _
								"left outer join OLKCustomSearchAlterNames T1 on T1.ObjectCode = T0.ObjectCode and T1.ID = T0.ID and T1.LanID = " & Session("LanID") & " " & _
								"where T0.ObjectCode = 2 and T0.Status = 'Y' and exists(select '' from OLKCustomSearchSession where ObjectCode = T0.ObjectCode and ID = T0.ID and SessionID = 'P') " & _
								"order by T0.Ordr "
						set rSearch = Server.CreateObject("ADODB.RecordSet")
						rSearch.open sql, conn, 3, 1
						If rSearch.recordcount > 0 Then %>
                	    <input type="button" value="<%=getsearchclientsLngStr("LtxtAdvance")%>" name="B2" onclick="javascipt:window.location.href='operaciones.asp?cmd=<% If rSearch.recordcount > 1 Then %>searchclient2<% Else %>adSearch&ID=<%=rSearch("ID")%>&adObjID=2<% End If %>'"><% End If %></td>
                    </tr>
                  </table>
                  </center>
                </div>
                </td>
              </tr>
            </table>
            <input type="hidden" name="D1" value="">
            <input type="hidden" name="D2" value="">
          </form>
          </td>
        </tr>
        <tr>
          <td width="100%">&nbsp;</td>
        </tr>
      </table>
      </td>
    </tr>
    </table>
  </center>
</div>
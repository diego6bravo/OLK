<% addLngPathStr = "inv/" %>
<!--#include file="lang/delSearchOrder.asp" -->
<%
If Request("bodega") <> "" Then Session("bodega") = Request("bodega")
If Request("Type") <> "" Then Session("Type") = Request("Type")
%>
<script Language="JavaScript">
function valFrm()
{
	ObjCode = document.frmOrder.ObjCode;
	if (ObjCode.length)
	{
		var found = false;
		for (var i = 0;i<ObjCode.length;i++)
		{
			if (ObjCode[i].checked)
			{
				found = true;
				break;
			}
		}
		if (!found)
		{
			alert('<%=getdelSearchOrderLngStr("LtxtSelDocType")%>');
			ObjCode[0].focus();
			return false;
		}
	}
	else
	{
		if (!ObjCode.checked)
		{
			alert('<%=getdelSearchOrderLngStr("LtxtSelDocType")%>');
			ObjCode.focus();
			return false;
		}
	}
	
	if (document.frmOrder.txtOrderNum.value == '')
	{
		alert('<%=getdelSearchOrderLngStr("LtxtValOrderNum")%>');
		document.frmOrder.txtOrderNum.focus();
		return false;
	}
	else if (!MyIsNumeric(document.frmOrder.txtOrderNum.value))
	{
		alert('<%=getdelSearchOrderLngStr("DtxtValNumVal")%>');
		document.frmOrder.txtOrderNum.focus();
		return false;
	}
	return true;
}
</script>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
  <form action="operaciones.asp" method="post" onsubmit="javascript:return valFrm();" name="frmOrder">
    <tr>
      <td bgcolor="#9BC4FF">

      <table border="0" cellpadding="0" bordercolor="#111111" width="100%" id="AutoNumber1">
        <tr>
          <td width="100%">
          <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><!--#include file="delOrderTitle.asp"-->
          </font></b></td>
        </tr>
        <tr>
          <td width="100%">
          <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><%=getdelSearchOrderLngStr("LttlSearchSourceDoc")%>
          </font></b></td>
        </tr>
        <tr>
          <td width="100%">
          <p align="center"><font size="1" face="Verdana">
          <table cellpadding="0" cellspacing="0" border="0">
          <%
          sql = "select T0.ObjectCode,  " & _
				"Case T0.ObjectCode " & _
				"	When 13 Then (select Singular from OLKAlterNames where LanID = " & Session("LanID") & " and AlterID = 5) " & _
				"	When 15 Then (select Singular from OLKAlterNames where LanID = " & Session("LanID") & " and AlterID = 11) " & _
				"	When 16 Then (select Singular from OLKAlterNames where LanID = " & Session("LanID") & " and AlterID = 12) " & _
				"	When 17 Then (select Singular from OLKAlterNames where LanID = " & Session("LanID") & " and AlterID = 7) " & _
				"	When 18 Then (select Singular from OLKAlterNames where LanID = " & Session("LanID") & " and AlterID = 17) " & _
				"	When 20 Then (select Singular from OLKAlterNames where LanID = " & Session("LanID") & " and AlterID = 15) " & _
				"	When 21 Then (select Singular from OLKAlterNames where LanID = " & Session("LanID") & " and AlterID = 16) " & _
				"	When 22 Then (select Singular from OLKAlterNames where LanID = " & Session("LanID") & " and AlterID = 14) " & _
				"End ObjectDesc, Case When ObjectCode in (17,22) Then 'Y' Else 'N' End Checked " & _
				"from OLKInOutSettings T0  " & _
				"where T0.Type = '" & Left(Request("Type"), 1) & "' "
				
			If Session("useraccess") = "U" Then
				sql = sql & " and ObjectCode in (select ObjectCode from OLKCommon..OLKAuthorization where AutID in (" & myAut.GetInOutAuthorization(Left(Request("Type"), 1)) & ")) "
			End If			
							
			sql = sql & "order by T0.DocType, T0.ObjectCode"
			set rs = conn.execute(sql)
			do while not rs.eof %>
				<tr>
					<td><input type="radio" name="ObjCode" <% If rs("Checked") = "Y" Then %>checked<% End If %> id="ObjCode<%=rs("ObjectCode")%>" value="<%=rs("ObjectCode")%>"><label for="ObjCode<%=rs("ObjectCode")%>"><font size="1" face="Verdana"><%=rs("ObjectDesc")%><% If Request("Type") = "O" and rs("ObjectCode") = 13 Then %>&nbsp;(<%=getdelSearchOrderLngStr("LtxtReserved")%>)<% End If %></font></label></td>
				</tr>
			<% rs.movenext
			loop %>
			</table>
          </font></td>
        </tr>
        <tr>
          <td width="100%">
            <table border="0" cellpadding="0" cellspacing="1"  bordercolor="#111111" width="100%" id="AutoNumber2">
              <tr>
                <td width="100%">
                <div align="center">
                  <center>
                  <table border="0" cellpadding="0" cellspacing="1" style="border-collapse: collapse" bordercolor="#111111" width="100" id="AutoNumber3">
                    <tr>
                      <td width="100%">
       				 <input type="number" name="txtOrderNum" size="20"></td>
                    </tr>
                    <tr>
                      <td width="100%">
                      <p align="center">
        			<input type="submit" name="btnSearch" value="<%=getdelSearchOrderLngStr("DbtnSearch")%>"></td>
                    </tr>
                  </table>
                  </center>
                </div>
            
                </td>
              </tr>
              </table>
          </td>
        </tr>
        </table>
      </td>
    </tr>
    <input type="hidden" name="cmd" value="invChkInOutSearch">
    <input type="hidden" name="Type" value="<%=Request("Type")%>">
    </form>
    </table>
  </center>
</div>
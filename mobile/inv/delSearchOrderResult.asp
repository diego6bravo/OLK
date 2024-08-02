<% addLngPathStr = "inv/"
If Request("ObjCode") <> "" Then Session("ObjCode") = Request("ObjCode")
 %>
<!--#include file="lang/delSearchOrderResult.asp" -->
<%
If Request("reProcess") = "Y" Then CreateIOData
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "DBOLKGetInOutData" & Session("ID")
cmd.Parameters.Refresh()
cmd("@LanID") = Session("LanID")
cmd("@ObjectCode") = Session("ObjCode")
cmd("@Type") = Session("Type")
cmd("@DocNum") = Request("txtOrderNum")
cmd("@WhsCode") = Session("bodega")
set rs = cmd.execute()

 %>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
    <tr>
      <td bgcolor="#9BC4FF">
      <table border="0" cellpadding="0" bordercolor="#111111" width="100%" id="AutoNumber1">
      	<% If Not rs.eof Then
      	If Not IsNull(rs("LogNum")) Then Session("IORetVal") = rs("LogNum") %>
        <tr>
          <td width="100%">
          <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1">
          <%=getdelSearchOrderResultLngStr("LttlSearchResult")%>&nbsp;-&nbsp;<%=rs("ObjDesc")%>
          </font></b></td>
        </tr>
		  <% 
		  If rs("DocStatus") = "C" Then %>
        <tr>
          <td width="100%" style="text-align: justify; ">
			<font face="Verdana" size="1"><% 
			varText = getdelSearchOrderResultLngStr("LtxtDocStatusClosed")
			varText = Replace(varText, "{0}", LCase(rs("TargetDesc")))
			varText = Replace(varText, "{1}", LCase(rs("ObjDesc")))
			varText = Replace(varText, "{2}", rs("ObjDesc"))
			varText = Replace(varText, "{3}", Request("txtOrderNum"))
			Response.Write varText %></font></td>
        </tr>
        <% ElseIf rs("DocType") = "S" Then %>
        <tr>
          <td width="100%" style="text-align: justify; ">
			<font face="Verdana" size="1"><% 
			varText = getdelSearchOrderResultLngStr("LtxtDocTypeService")
			varText = Replace(varText, "{0}", LCase(rs("TargetDesc")))
			varText = Replace(varText, "{1}", LCase(rs("ObjDesc")))
			varText = Replace(varText, "{2}", rs("ObjDesc"))
			varText = Replace(varText, "{3}", Request("txtOrderNum"))
			Response.Write varText %></font></td>
        </tr>
          <% ElseIf rs("CheckStatus") = "S" Then %>
          <tr>
          	<td>
          		<table cellpadding="0" cellspacing="0" border="0" width="100%">
					<tr>
						<td colspan="2" style="text-align: justify; ">
						<font face="Verdana" size="1"><% 
						varText = getdelSearchOrderResultLngStr("LtxtConfReProcess")
						varText = Replace(varText, "{0}", rs("ObjDesc"))
						varText = Replace(varText, "{1}", Request("txtOrderNum"))
						Response.Write varText %></font>
						</td>
					</tr>
					<tr>
						<td colspan="2" height="5px"></td>
					</tr>
					<tr>
						<td align="center"><input type="button" name="btnConfirm" value="<%=getdelSearchOrderResultLngStr("DtxtYes")%>" style="width: 80px;" onclick="javascript:doReProcess()"></td>
						<td align="center"><input type="button" name="btnCancel" value="<%=getdelSearchOrderResultLngStr("DtxtNo")%>" onclick="javascript:window.location.href='?cmd=invChkInOut'" style="width: 80px;"></td>
					</tr>
					<tr>
						<td colspan="2" height="5px"></td>
					</tr>
				</table>
          	</td>
          </tr>
          <% ElseIf Session("ObjCode") = "18" and rs("IsIns") <> "Y" and Session("Type") = "I" Then %>
        <tr>
          <td width="100%" style="text-align: justify; ">
			<font face="Verdana" size="1"><% 
			varText = ""
			varText = getdelSearchOrderResultLngStr("LtxtInvDocStatusIsIns")
			varText = Replace(varText, "{0}", Request("txtOrderNum"))
			Response.Write varText %></font></td>
        </tr>
          <% ElseIf Session("ObjCode") = "13" and rs("IsIns") <> "Y" and Session("Type") = "O" Then %>
        <tr>
          <td width="100%" style="text-align: justify; ">
			<font face="Verdana" size="1"><% 
			varText = ""
			varText = getdelSearchOrderResultLngStr("LtxtInvDocStOutIsIns")
			varText = Replace(varText, "{0}", Request("txtOrderNum"))
			Response.Write varText %></font></td>
        </tr>
          <% ElseIf rs("VerfyClosedLines") = "Y" Then %>
          <tr>
          	<td>
          		<table cellpadding="0" cellspacing="0" border="0" width="100%">
					<tr>
						<td colspan="2" style="text-align: justify; ">
						<font face="Verdana" size="1">
						<%=Replace(getdelSearchOrderResultLngStr("LtxtConfirmClosedLine"), "{0}", Request("txtOrderNum"))%></font>
						</td>
					</tr>
					<tr>
						<td colspan="2" height="5px"></td>
					</tr>
					<tr>
						<td align="center"><input type="button" name="btnConfirm" value="<%=getdelSearchOrderResultLngStr("DtxtYes")%>" style="width: 80px;" onclick="javascript:window.location.href='?cmd=invChkInOutCheck&txtOrderNum=<%=Request("txtOrderNum")%><% If IsNull(rs("CheckStatus")) Then %>&CreateStatus=Y<% End If %>'"></td>
						<td align="center"><input type="button" name="btnCancel" value="<%=getdelSearchOrderResultLngStr("DtxtNo")%>" onclick="javascript:window.location.href='?cmd=invChkInOut'" style="width: 80px;"></td>
					</tr>
					<tr>
						<td colspan="2" height="5px"></td>
					</tr>
				</table>
          	</td>
          </tr>
          <% ElseIf Session("ObjCode") = "17" and rs("Confirmed") <> "Y"  Then %>
        <tr>
          <td width="100%" style="text-align: justify; ">
			<font face="Verdana" size="1"><%=Replace(getdelSearchOrderResultLngStr("LtxtOrdrAut"), "{0}", Request("txtOrderNum"))%></font></td>
        </tr>
          <% Else
          	If IsNull(rs("CheckStatus")) Then
          		CreateIOData
          	End If
          	Response.Redirect "?cmd=invChkInOutCheck&txtOrderNum=" & Request("txtOrderNum") %>
		  <% End If %>
 		 <% Else %>
        <tr>
          <td width="100%">
			<p align="center"><font face="Verdana" size="1"><%=getdelSearchOrderResultLngStr("DtxtNoData")%></font></td>
        </tr>
  		<% End If %>
       </table>
       </td>
   </tr>
   </table>
  </center>
</div>
<script language="javascript">
function doReProcess()
{
	document.frmResubmit.submit();
}
</script>
<form name="frmResubmit" method="post" action="operaciones.asp">
<% For each itm in Request.Form %>
<input type="hidden" name="<%=itm%>" value="<%=Request.Form(itm)%>">
<% Next %>
<input type="hidden" name="reProcess" value="Y">
</form>
<% Sub CreateIOData
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "DBOLKCreateIOData" & Session("ID")
cmd.Parameters.Refresh()
cmd("@ObjectCode") = Session("ObjCode")
cmd("@Type") = Session("Type")
cmd("@DocNum") = Request("txtOrderNum")
cmd.execute()

Session("IORetVal") = cmd("@LogNum")
End Sub %>
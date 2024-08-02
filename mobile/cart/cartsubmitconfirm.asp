<% addLngPathStr = "cart/" %>
<!--#include file="lang/cartsubmitconfirm.asp" -->
<%
If Request("s") <> "E" Then
	sql = "EXEC OLKCommon..DBolkmensaje" & Session("ID") & " @lognum = " & Session("ConfRetVal") & ", @LanID = " & Session("LanID")
	set rs = conn.execute(sql)
Else
	sql = "select ErrMessage from R3_ObsCommon..TLOG where LogNum = " & Session("RetVal")
	set rs = conn.execute(sql)
End If
%><div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
    <tr>
      <td bgcolor="#9BC4FF">
      <table border="0" cellpadding="0" bordercolor="#111111" width="100%" id="AutoNumber1">
        <tr>
          <td width="100%" bgcolor="#75ACFF">
          <p align="center">
      <b>
      <font face="Verdana" size="1"><%=getcartsubmitconfirmLngStr("LtxtShopCartConf")%></font></b></p>
          </td>
        </tr>
        <tr>
          <td width="100%">&nbsp;</td>
         </tr>
        <% If Request("s") = "E" Then %> 
        <tr>
          <td width="100%" align="center">
			<img src="images/errorIcon.gif">
		  </td>
        </tr>      
        <% End If %>
        <tr>
          <td width="100%">
          <p align="center"><font size="1" face="Verdana"><% 
          Select Case Request("s")
          	Case "S" %>
          	<b><%=Replace(Replace(getcartsubmitconfirmLngStr("LtxtDocAddOK"), "{0}", rs("DocName")), "{1}", rs("DocNum"))%></b>
	     <% Case "H" %>
	     	<b><%=Replace(getcartsubmitconfirmLngStr("LtxtWaitForConf"), "{0}", Session("ConfRetVal"))%></b>
	     <% Case "E" %>
	        <b><%=getcartsubmitconfirmLngStr("LtxtErrAddDoc")%></b><br><%=rs("ErrMessage")%>
         <% End Select %></font></td>
        </tr>
        <% If Request("s") <> "E" Then
        objCode = rs("Object")
		hasPrint = rs("HasPrint") = "Y"
		%>
        <form name="frmViewDoc" method="post" action="cxcDocDetail.asp">
        <tr>
          <td width="100%" align="center">
			<input type="image" src="images/print_OLK.gif" border="0" align="middle" id="fp1"><b><font size="1" face="Verdana"><label for="fp1"><%=getcartsubmitconfirmLngStr("DtxtPrint")%></label></font></b></td>
        </tr>
        <input type="hidden" name="DocType" value="<% If rs("Status") = "S" Then %><%=rs("object")%><% Else %>-2<% End If %>">
        <input type="hidden" name="DocEntry" value="<% If rs("Status") = "S" Then %><%=rs("ObjectCode")%><% Else %><%=Session("ConfRetVal")%><% End IF %>">
        </form>
		<% If hasPrint Then
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKGetObjectPrint" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@ObjCode") = objCode
		cmd("@UserType") = "P"
		cmd("@LanID") = Session("LanID")
		set rp = Server.CreateObject("ADODB.RecordSet")
		set rp = cmd.execute()
		do while not rp.eof
		secID = rp("SecID") %>
		<tr class="FirmTbl" id="trSubmitPrint">
			<td>
			<p align="center"><input type="button" name="btnPrint<%=secID%>" id="btnPrint<%=secID%>" value="<%=rp("SecName")%>" onclick="javascript:openPrint(<%=secID%>, '<%=rp("LinkData")%>');"></td>
		</tr>
		<% rp.movenext
		loop
		set rp = nothing 
		End If %>
        <% Else %>
        <tr>
          <td width="100%" align="center">
			<input type="button" name="btnRestore" value="<%=getcartsubmitconfirmLngStr("DtxtRestore")%>" onclick="javascript:window.location.href='operaciones.asp?cmd=cart'">&nbsp;-&nbsp;
			<input type="button" name="btnRetry" value="<%=getcartsubmitconfirmLngStr("DtxtRetry")%>" onclick="javascript:window.location.href='operaciones.asp?cmd=cartSubmit&retry=Y'">
		  </td>
        </tr>     
        <% End If %>
      </table>
      </td>
    </tr>
    </table>
  </center>
</div>
<% If hasPrint Then %>
<script type="text/javascript">
function openPrint(secID, linkData)
{
	OpenWin = window.open('sectionClean.asp?secID=' + secID + '&' + linkData.replace('{0}', '<%=rs("DocNum")%>'), 'Print', 'toolbar=no,menubar=no,location=no,scrollbars=yes,resizable=yes, width=760,height=540');
}
</script>
<%
End If
If Session("NotifyAdd") Then
	Session("NotifyAdd") = False
	sql = "EXEC OLKCommon..DBOLKObjAlert" & Session("ID") & " " & Session("ConfRetVal") & ", " & Session("branch") & ", 'V', '" & getMyLng & "'"
	conn.execute(sql)
End If
	
'Session("RetVal") = ""
%>
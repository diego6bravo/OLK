<!--#include file="lang/listpen.asp" -->
<%
set rs = server.createobject("ADODB.RecordSet")

If Request("delLog") <> "" Then
	sql = "update R3_ObsCommon..TLOG set Status = 'B' where LogNum = " & Request("delLog")
	conn.execute(sql)
End If

ObjCode = ""
If myApp.EnableOQUT Then ObjCode = "23"
If myApp.EnableORDR Then ObjCode = myAut.ConcValue(ObjCode, "17")
If myApp.EnableODLN Then ObjCode = myAut.ConcValue(ObjCode, "15")
If myApp.EnableODPIReq Then ObjCode = myAut.ConcValue(ObjCode, "203")
If myApp.EnableODPIInv Then ObjCode = myAut.ConcValue(ObjCode, "204")
If myApp.EnableOINV or myApp.EnableOINVRes Then ObjCode = myAut.ConcValue(ObjCode, "13")

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKSeachOpenDocs" & Session("ID")
cmd.Parameters.Refresh()
cmd("@AllowAgentAccessCDoc") = GetYN(myApp.AllowAgentAccessCDoc)
cmd("@LanID") = Session("LanID")
cmd("@SlpCode") = Session("vendid")

If Request("cmd") = "activeClient" Then 
	If Request("CardCodeFrom") <> "" Then cmd("@CardCodeFrom") = Session("UserName")
	If Request("CardCodeTo") <> "" Then cmd("@CardCodeTo") = Session("UserName")
Else
	If Request("CardCodeFrom") <> "" Then cmd("@CardCodeFrom") = Request("CardCodeFrom")
	If Request("CardCodeTo") <> "" Then cmd("@CardCodeTo") = Request("CardCodeTo")
End If

If Request("ItemCodeFrom") <> "" Then cmd("@ItemCodeFrom") = Request("ItemCodeFrom")
If Request("ItemCodeTo") <> "" Then cmd("@ItemCodeTo") = Request("ItemCodeTo")

If Request("LogNumFrom") <> "" Then cmd("@LogNumFrom") = Request("LogNumFrom")
If Request("LogNumTo") <> "" Then cmd("@LogNumTo") = Request("LogNumTo") 

cmd("@All") = "Y"

If AsignedSlp or not myAut.HasAuthorization(97) Then 
	cmd("@All") = "N"
End If

If Request("Comments") <> "" Then cmd("@Comments") = Request("Comments")

If Request("GroupNameFrom") <> "" Then cmd("@GroupNameFrom") = Request("GroupNameFrom")
If Request("GroupNameTo") <> "" Then cmd("@GroupNameTo") = Request("GroupNameTo")

If Request("CountryFrom") <> "" Then cmd("@CountryNameFrom") = Request("CountryFrom")
If Request("CountryTo") <> "" Then cmd("@CountryNameTo") = Request("CountryTo")

If Request("CardString") <> "" Then cmd("@CardString") = Request("CardString")

If Request("DocType") <> "" Then cmd("@DocType") = Request("DocType")

If Request("SlpCodeFrom") <> "" Then cmd("@SlpNameFrom") = Request("SlpCodeFrom")
If Request("SlpCodeTo") <> "" Then cmd("@SlpNameTo") = Request("SlpCodeTo")

If Request("dtFrom") <> "" Then cmd("@DateFrom") = SaveCmdDate(Request("dtFrom"))
If Request("dtTo") <> "" Then cmd("@DateTo") = SaveCmdDate(Request("dtTo"))

If Request("orden1") <> "" Then cmd("@Order") = Request("orden1")
If Request("orden2") <> "" Then cmd("@OrderDir") = Request("orden2")
cmd("@Objects") = ObjCode
cmd("@EnableCashInv") = "N"
cmd("@EnableInvRes") = GetYN(myApp.EnableOINVRes)
cmd("@EnableInv") = GetYN(EnableInv)

rs.CursorLocation = 3 ' adUseClient
rs.open cmd
rs.PageSize = 10
rs.CacheSize = 10

If Request("p") <> "" Then iCurPage = CInt(Request("p")) Else iCurPage = 1
iPageCount = rs.PageCount
%>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
    <tr>
      <td>
      <img src="images/spacer.gif" width="100%" height="1" border="0" alt></td>
    </tr>
    <tr>
      <td bgcolor="#9BC4FF">
      <table border="0" cellpadding="0" bordercolor="#111111" width="100%" id="AutoNumber1">
        <tr>
          <td width="100%">
          <table style="width: 100%">
			<tr>
				<td><input type="button" name="btnSearch" value="<%=getlistpenLngStr("DtxtSearch")%>" onclick="javascript:document.frmGo.cmd.value='searchPend';document.frmGo.submit();"></td>
				<td align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><%=getlistpenLngStr("LtxtOpenDocs")%></font></b></td>
			</tr>
			</table>
          </td>
        </tr>
		<% If Not rs.Eof Then
		rs.AbsolutePage = iCurPage
		LogNum = ""
		For i = 1 to rs.PageSize
			If i > 1 Then LogNum = LogNum & ", "
			LogNum = LogNum & rs("LogNum")
			rs.movenext
			If rs.eof then exit for
		Next
		rs.close
		
		sql = 	"select T0.LogNum, Convert(int,DocDate), DocDate, Object, CardCode, ReserveInvoice, " & _
		"OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRD', 'CardName', T1.CardCode, T1.CardName) CardName, DocCur, Comments, OLKCommon.dbo.DBOLKDocTotal" & Session("ID") & "(T0.LogNum) DocTotal, T0.Status " & _
		"from r3_obscommon..tlog T0 " & _
		"inner join r3_obscommon..tdoc T1 on T1.LogNum = T0.LogNum " & _
		"where T0.LogNum in (" & LogNum & ") " & VendId & " " & AgentClientsFilter & _
		"order by 2 desc "
		
		set rs = conn.execute(sql)
		
		do while not rs.eof
			Enable = True
			Select Case rs("Object")
		  		Case 17
		 			If Not myApp.EnableORDR Then Enable = False
		  		Case 23
		  			If Not myApp.EnableOQUT Then Enable = False
		  	End Select %>
        <tr>
          <td width="100%">
          <table border="0" cellpadding="0" cellspacing="1"  bordercolor="#111111" width="100%" id="AutoNumber2">
            <tr>
              <td width="4%" bgcolor="#66A4FF">
    <a href="javascript:<% If Enable Then %>doGoDoc(<%=rs("Object")%>, '<%=rs("LogNum")%>', '<%=Replace(myHTMLEncode(rs("CardCode")), "'", "\'")%>', '<%=rs("status")%>');<% Else %>listPendAlert(<%=rs("Object")%>);<% End If %>">
    <img border="0" src="images/<%=Session("rtl")%>flechaselec.gif" align="left"></a></td>
              <td width="28%" bgcolor="#66A4FF"><font face="verdana" color="#000000" size="1"><%=RS("LogNum")%></font></td>
              <td width="38%" bgcolor="#66A4FF"><font face="verdana" color="#000000" size="1"><%=FormatDate(RS("DocDate"), True)%></font></td>
              <td width="30%" bgcolor="#66A4FF"><font face="verdana" color="#000000" size="1"><%
		     Select Case RS("Object")
			     Case 17
			     	Response.write txtOrdr
			     Case 23
			     	Response.write txtQuote
			     Case 24
			     	Response.write txtRct
			     Case 13
			     	Select Case rs("ReserveInvoice")
			     		Case "Y"
					     	Response.write txtInvRes
					     Case Else
					     	Response.write txtInv
					End Select
			     Case 15
			     	Response.write txtOdln
		     End Select %></font></td>
            </tr>
            <tr>
              <td width="4%">
    <a href="javascript:delLogNum(<%=RS("LogNum")%>);">
    <img border="0" src="images/remove.gif"></a></td>
              <td width="28%"><font size="1" face="verdana" color="#000000"><a href="operaciones.asp?cmd=datos&card=<%=CleanItem(myHTMLEncode(RS("CardCode")))%>"><font color="#000000"><%=RS("CardCode")%></font></a></td>
              <td width="38%"><font size="1" face="verdana" color="#000000"><%=RS("CardName")%></font></td>
              <td width="30%" align="right"><p align="right" dir="ltr"><font size="1" face="verdana" color="#000000"><nobr><%=RS("DocCur")%>&nbsp;<%=FormatNumber(RS("DocTotal"), myApp.SumDec)%></nobr></font></p></td>
            </tr>
            <% If rs("Comments") <> "" Then %>
            <tr>
              <td width="100%" colspan="4"><b><font size="1" face="verdana" color="#000000">
				<%=getlistpenLngStr("DtxtObservations")%>: <%=rs("Comments")%></font></b>
			</td>
            </tr>
  			<% End If %>
            </table>
          </td>
        </tr>
     <% 
     rs.movenext
	 loop %>
        <tr>
          <td width="100%" colspan="4">
			<table border="0" cellpadding="0" cellspacing="1" width="100%" id="table1" dir="ltr">
				<tr>
					<td width="16">
					<% If iCurPage > 1 Then %><a href="javascript:goP(<%=iCurPage-1%>);"><img border="0" src="images/flecha_prev.gif" width="16" height="16"></a><% Else %>&nbsp;<% End If %></td>
					<td>
					<p align="center">
					<select name="cmbPage" size="1" onchange="javascript:javascript:goP(this.value);">
					<% For i = 1 to iPageCount %>
					<option value="<%=i%>" <% If i = iCurPage Then %>selected<% End If %>><%=i%></option>
					<% Next %>
					</select></td>
					<td width="16">
					<% If iCurPage < iPageCount Then %><a href="javascript:goP(<%=iCurPage+1%>);"><img border="0" src="images/flecha_next.gif" width="16" height="16"></a><% End If %></td>
				</tr>
			</table>
			</td>
        </tr>
        <% Else %>
        <tr>
          <td width="100%" align="center"><b>
			<font size="1" face="verdana" color="#000000"><%=getlistpenLngStr("DtxtNoData")%></font></b></td>
        </tr>
        <% End If %>
        <tr>
          <td width="100%"><hr color="#3385FF" size="1"></td>
        </tr>
        
        </table>
      </td>
    </tr>
    </table>
  </center>
</div>
<form name="frmGo" method="post" action="operaciones.asp">
<% For each itm in Request.Form
If itm <> "p" and itm <> "delLog" Then %>
<input type="hidden" name="<%=itm%>" value="<%=Request.Form(itm)%>">
<% End If 
Next %>
<% For each itm in Request.QueryString
If itm <> "p" Then %>
<input type="hidden" name="<%=itm%>" value="<%=Request.QueryString(itm)%>">
<% End If 
Next %>
<input type="hidden" name="p" value="<%=iCurPage%>">
<input type="hidden" name="delLog" value="">
</form>
<script language="javascript">
function goP(p)
{
	document.frmGo.cmd.value = 'pendientes';
	document.frmGo.p.value = p;
	document.frmGo.delLog.value = '';
	document.frmGo.submit();
}
function delLogNum(lognum)
{
	if (confirm('<%=getlistpenLngStr("LtxtConfDelDoc")%>'.replace('{0}', lognum)))
	{
		document.frmGo.cmd.value = 'pendientes';
		document.frmGo.delLog.value = lognum;
		document.frmGo.submit();
	}
}
function listPendAlert(obj) {
var objType;
switch (obj) {
	case 15:
		objType = "<%=txtOdln%>";
		break;
	case 17:
		objType = "<%=txtOrdr%>";
		break;
	case 23:
		objType = "<%=txtQuote%>";
		break;
	case 24:
		objType = "<%=txtRct%>";
		break;
	case 48:
		objType = "<%=txtInv%>/<%=txtRct%>";
		break;
	case 13:
		objType = "<%=txtInv%>";
		break;
	case 4:
		objType = "<%=getlistpenLngStr("DtxtItem")%>";
		break;
	case 2:
		objType= "<%=txtClient%>"
		break;
}
alert('<%=getlistpenLngStr("LtxtDisObj")%>'.replace('{0}', objType));
}


function confReopen()
{
	return confirm('<%=getlistpenLngStr("LtxtConfReOpen")%>')
}

function doGoDoc(obj, logNum, CardCode, Status)
{
	if (Status == 'H') if (!confReopen()) return;
	
	document.frmGoDoc.doc.value = logNum;
	document.frmGoDoc.cl.value = CardCode;
	document.frmGoDoc.status.value = Status;
	
	document.frmGoDoc.submit();
}
</script>
<form name="frmGoDoc" action="go.asp" method="post">
<input type="hidden" name="doc" value="">
<input type="hidden" name="cl" value="">
<input type="hidden" name="status" value="">
</form>
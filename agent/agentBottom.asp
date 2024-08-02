<% addLngPathStr = "" %>
<!--#include file="lang/agentBottom.asp" -->
</div></td>
      </tr>
      </table>
    </td>
  </tr>
  <tr id="pageBottom">
    <td colspan="2" background="ventas/images/art_ventas1_r5_c1.jpg" style="padding-top: 22px;"><p dir="ltr"><font face="Tahoma" size="1" color="#FFFFFF">&nbsp;&nbsp;v<%=OLKVersion%>&nbsp;(SBO <%=myApp.VSystem%>)</font></p>
    </td>
    <td colspan="4" background="ventas/images/art_ventas1_r5_c7.jpg" valign="top">
    <img border="0" src="ventas/images/<%=Session("rtl")%>art_conimg_r3_c1_down.jpg" width="32" height="20"></td>
    <td colspan="7" background="ventas/images/art_ventas1_r5_c7.jpg" style="padding-top: 22px; height: 50px;">
      <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><font size="1" color="#FFFFFF">Copyright &#169; 
		2002-2012 TopManage &reg;&nbsp; - <%=getagentBottomLngStr("DtxtEMail")%>: </font>
		<font face="Tahoma" color="#c0c0c0" size="1">
		<a href="mailto:info@topmanage.com.pa"><font color="#FFFFFF" size="1">info@topmanage.com.pa</font></a></font><font size="1" color="#FFFFFF"> 
		- <%=getagentBottomLngStr("DtxtPhone")%>: 507.300.7200&nbsp; </font>
    	<font face="Tahoma" size="1" color="#FFFFFF">&nbsp;&nbsp; </font></p>
    </td>
  </tr>
</table>
<script language="javascript">function goOp(CardCode) { document.frmGoOp.cCode.value = CardCode; document.frmGoOp.submit(); }</script>
<form method="post" action="activeClient.asp" name="frmGoOp">
<input type="hidden" name="cCode" value="">
<input type="hidden" name="cmd" value="activeClient">
</form>
<form name="frmGoRep" id="frmGoRep" method="post" action="portal/viewRepVals.asp" target="RepVals">
<input type="hidden" name="rsIndex" value="">
<input type="hidden" name="pop" value="">
<input type="hidden" name="AddPath" value="">
<input type="hidden" name="cmd" value="report"></form>
<script>
if (!document.layers)
document.write('<div id="divStayTop" style="position:absolute">')
</script>
<layer id="divStayTop">
<table cellpadding="0" cellspacing="0" width="117" border="0">
<% If Not Session("Touch") Then %>
<tr>
	<td><a href="#pageTop"><img border="0" src="images/img_page_top.gif" alt="<%=getagentBottomLngStr("LtxtGoTop")%>"></a></td>
</tr>
<% End If %>
<% If Menu Then %>
<tr>
	<td>
	<table border="0" cellpadding="0" cellspacing="0" width="108">
		<tr>
			<td>
			<% Select Case searchCmd
			Case "clientsSearch" %><!--#include file="searchInc/searchClients.asp" -->
			<% Case "searchCart" %><!--#include file="searchInc/searchCartInc.asp" -->
			<% Case "searchCatalog" %><!--#include file="searchInc/searchCatalog.asp" -->
			<% Case "report" %><!--#include file="searchInc/repReload.asp" -->
			<% Case "docsConfirmation" %><!--#include file="searchInc/confirmReload.asp"-->
			<% Case "extPollview" %><!--#include file="searchInc/reload.asp"-->
			<% end select %>
			<% If SearchCmd = "searchOfertsX" or SearchCmd = "searchItemX" or SearchCmd = "searchCardX" or SearchCmd = "searchDocX" or SearchCmd = "searchActX" or SearchCmd = "searchSOX" Then %>
			<!--#include file="searchInc/backToSearch.asp" --><% End If %></td>
		</tr>
	</table>
	</td>
</tr>
<% Else %>
<tr>
	<td style="height: 50px;"></td>
</tr>
<% end if %>
<% If Not Session("Touch") Then %>
<tr>
	<td><a href="#pageBottom"><br><img src="images/img_page_bottom.gif" border="0" alt="<%=getagentBottomLngStr("LtxtGoBottom")%>"></a></td>
</tr>
<% End If %>
</table>

<% If Menu and CartInfo Then %>
<div style="background-color: #0066ca; width: 106px;">
<% 
set rx = Server.CreateObject("ADODB.RecordSet")
sql = "select T0.RowType+Convert(nvarchar(20),T0.LineIndex) RowID, T0.RowQuery, IsNull(T1.AlterRowName, T0.RowName) RowName, T0.Align " & _
"from OLKCMREP T0 " & _
"left outer join OLKCMREPAlterNames T1 on T1.RowType = T0.RowType and T1.LineIndex = T0.LineIndex and T1.LanID = " & Session("LanID") & " " & _
"where T0.RowActive = 'Y' and T0.ShowV = 'Y' and RowQuery is not null " & _
"order by T0.RowOrder asc"
rx.open sql, conn, 3, 1
if not rx.eof then
sql = "declare @LogNum int set @LogNum = " & Session("RetVal") & " " & _
"declare @CardCode nvarchar(15) set @CardCode = N'" & saveHTMLDecode(Session("username"), False) & "' " & _
"declare @LanID int set @LanID = " & Session("LanID") & " " & _
"select "
do while not rx.eof
	If rx.bookmark > 1 Then sql = sql & ", "
	sql = sql & "(" & rx("RowQuery") & ") As '" & rx("RowID") & "{S}" & Replace(rx("RowName"), "'", "''") & "{S}" & rx("Align") &  "'"
rx.movenext
loop
sql = QueryFunctions(sql)
set rx = conn.execute(sql) %>
<div align="center"><b><font size="1" face="Verdana" color="#FFFFFF"><%=txtBasketMinRep%></font></b></div>
<div id="trMinRep" style="display: none;">
	<% If Not Session("Touch") Then %><div align="center" onmousedown="minRepMove('U');" id="minRepMoveUp" onmouseover="document.getElementById('btnMinRepUp').src='images/minRepScrollUp.jpg';" onmouseup="stopMinRepMove();" onmouseout="stopMinRepMove();document.getElementById('btnMinRepUp').src='images/minRepScrollUpW.jpg';"><img src="images/minRepScrollUpW.jpg" id="btnMinRepUp" alt=""></div><% End If %>
	<div style="border: 1px solid #FFFFFF;">
		<div id="scrollMinRep3" style="width:100%;height:200px;overflow:<% If Not Session("Touch") Then %>hidden<% Else %>scroll<% End If %>; position: relative;">
			<div style="width: 104px;  position: absolute; top: 0px; left: 0px; width: 104px;" id="tblMinRep">
			<% 
			For each fld in rx.Fields
			arrValues = Split(fld.Name, "{S}")
			rowID = arrValues(0)
			rowName = arrValues(1)
			rowAlign = ""
			Select Case arrValues(2)
			Case "L"
				rowAlign = "left"
			Case "C"
				rowAlign = "center"
			Case "R"
				rowAlign = "right"
			End Select %>
			<div class="MinRepTtl" align="center"><%=rowName%>&nbsp;</div>
			<div align="<%=rowAlign%>" class="MinRepItm" id="mrVal<%=rowID%>"><% If Not IsNull(fld) Then %><%=fld%><% End If %>&nbsp;</div>
			<% Next %>
			</div>
			<div id="tblMinRepWait" class="Transparency" style="display: none; position: absolute; left: 0px; top: 0px; filter:alpha(opacity=60); background: rgb(0,88,177); height: 200px; width: 100%;">
			<img src="design/0/images/ajax-cartrep-loader.gif" id="imgMinRepWait" style="display: none; position: absolute; top: 80px; left: 30px;">
			</div>
		</div>
	</div>
	<% If Not Session("Touch") Then %><div align="center" onmousedown="minRepMove('D');" id="minRepMoveDown" onmouseover="document.getElementById('btnMinRepDown').src='images/minRepScrollDown.jpg';" onmouseup="stopMinRepMove();" onmouseout="stopMinRepMove();document.getElementById('btnMinRepDown').src='images/minRepScrollDownW.jpg';"><img src="images/minRepScrollDownW.jpg" id="btnMinRepDown" alt=""></div><% End If %>
</div>
<% End If %>
</div>
<% End If %>


<!--END OF EDIT-->

</layer>


<script type="text/javascript">
<% If Session("rtl") = "" Then %>
JSFX_FloatDiv("divStayTop", 5, 150).flt();
<% Else %>
JSFX_FloatDiv("divStayTop", -120, 150).flt();
<% End If %>

</script>
<form name="frmChangeLng" method="post" action="" method="post">
<% For each itm in Request.Form
If itm <> "newLng" Then %>
<input type="hidden" name="<%=itm%>" value="<%=myHTMLEncode(Request.Form(itm))%>">
<% End If
Next
For Each itm in Request.QueryString
If itm <> "newLng" Then %>
<input type="hidden" name="<%=itm%>" value="<%=myHTMLEncode(Request.QueryString(itm))%>">
<% End If 
Next %>
<input type="hidden" name="newLng" value="">
</form>
<script language="javascript">
function goNavQry(navIndex, CatType)
{
	document.frmNavQry.document.value = CatType;
	document.frmNavQry.navIndex.value = navIndex;
	document.frmNavQry.submit();
}
</script>
<form name="frmNavQry" action="search.asp" method="post">
<input type="hidden" name="cmd" value="search<% If session("RetVal") = "" Then %>Catalog<% Else %>Cart<% End If %>">
<input type="hidden" name="focus" value="frmSmallSearch.string">
<input type="hidden" name="orden1" value="<% If myApp.GetDefCatOrdr = "C" Then %>OITM.ItemCode<% Else %>ItemName<% End If %>">
<input type="hidden" name="orden2" value="asc">
<input type="hidden" name="document" value="">
<input type="hidden" name="navIndex" value="">
</form>
<table cellpadding="0" border="0" width="28" bgcolor="#0066cc" id="tblSmallExport" style="display: none; position: absolute; z-index: 1" onmouseover="expTblID=window.clearInterval(expTblID);" onmouseout="expTblID=window.setTimeout('clearSmallExpTbl()', 1000);">
	<% If excell Then %>
	<tr>
		<td height="29" align="center" id="tdExpExcell" onmouseover="setSmallExpTDBorder(this, true);" onmouseout="setSmallExpTDBorder(this, false);"><a href="<% If InStr(excellaction, "javascript:") = 0 Then %><%=excellaction%>?excell=Y<%=excellVals%><% Else %><%=excellaction%><% End If %>">
		<img src="ventas/images/iconos_ventas_export_excell.gif" border="0" alt="<%=getagentBottomLngStr("LtxtExpExcell")%>"></a></td>
	</tr>
	<% End If %>
	<% If pdfDoc Then %>
	<tr>
		<td height="29" align="center" id="tdExpPDF" onmouseover="setSmallExpTDBorder(this, true);" onmouseout="setSmallExpTDBorder(this, false);"><a href="<% If pdfAction = "" Then %><% If curCmd = "cartsubmitconfirm" Then %>javascript:saveInPwd('<%=Request("status")%>');<% Else %>javascript:printCat('Y');<% End If %><% Else %><% If InStr(pdfAction, "javascript:") = 0 Then %><%=pdfAction%>?pdf=Y<%=excellVals%><% Else %><%=pdfAction%><% End If %><% End If %>">
		<img src="ventas/images/iconos_ventas_export_pdf.gif" border="0" alt="<%=getagentBottomLngStr("DtxtExpPDF")%>"></a></td>
	</tr>
	<% End If %>
</table>
<!--#include file="messages/messageAlert.asp"-->
<!--#include file="linkForm.asp"-->
<%
set myFlow = new FlowControl
myFlow.GenerateFlow
%>
</body>
<script>setAllSize();</script>
<% 
conn.close
set rh = nothing 
set rm = nothing
set rs = nothing
set rx = nothing
set rd = nothing
set rf = nothing 
%></html>
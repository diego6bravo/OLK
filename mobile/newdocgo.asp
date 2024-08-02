<!--#include file="lang/newdocgo.asp" -->
<% varx = 0 %>
<SCRIPT LANGUAGE="JavaScript">

<!-- Begin
function Start(page) {
OpenWin = this.open(page, "CtrlWindow", "toolbar=no,menubar=no,location=no,scrollbars=no,resizable=no, height=378,width=420");
}
// End -->
</SCRIPT>

<%
sqlAdd = ""
sqlAdd2 = ""

If Session("useraccess") = "U" and (not myAut.HasAuthorization(60) or not myAut.HasAuthorization(97)) Then
	sqlAdd = " and T1.SlpCode = " & Session("vendid") & " "
	sqlAdd2 = " and tdoc.SLPCode = " & Session("vendid") & " "
End If

ObjCode = ""
If myApp.EnableOQUT Then ObjCode = "23"
If myApp.EnableORDR Then
	If ObjCode <> "" Then ObjCode = ObjCode & ", "
	ObjCode = ObjCode & "17"
End If
If myApp.EnableOINV Then
	If ObjCode <> "" Then ObjCode = ObjCode & ", "
	ObjCode = ObjCode & "13"
End If
If ObjCode = "" Then ObjCode = "-1"

sql = "declare @CardCode nvarchar(15) set @CardCode = N'" & saveHTMLDecode(Request("c1"), False) & "'  " & _
"declare @DocCount int " & _
"declare @TotalDocs numeric(19,6) " & _
"select Count(T0.lognum) DocCount,  " & _
"(select ISNULL(sum(Price * Quantity),0) As Total  " & _
"from r3_obscommon..doc1  " & _
"where lognum in (  " & _
"	select tlog.lognum from r3_obscommon..tlog tlog  " & _
"	inner join r3_obscommon..tdoc tdoc on tdoc.lognum = tlog.lognum where Company = '" & Session("olkDB") & "' and tlog.object in (" & ObjCode & ") " & sqlAdd2 & " and status in ('R', 'H') and cardcode = @CardCode and tlog.LogNum not in (select LogNum from OLKClientsDocControl where CardCode = @CardCode))) Total  " & _
"from r3_obscommon..tlog T0 " & _
"inner join r3_obscommon..tdoc T1 on T1.lognum = T0.lognum  " & _
"inner join OCRD T2 on T2.CardCode = T1.CardCode collate database_default " & _
"where Company = '" & Session("olkDB") & "' and T0.object in (" & ObjCode & ") " & sqlAdd & " and status in ('R', 'H') and T1.cardcode = @CardCode and T0.LogNum not in (select LogNum from OLKClientsDocControl where CardCode = @CardCode) "
set rs = conn.execute(sql)

DocCount = rs("DocCount")
DocTotal = rs("Total")

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetCardSpecData" & Session("ID")
cmd.Parameters.Refresh()
cmd("@DataType") = 1
cmd("@LanID") = Session("LanID")
cmd("@CardCode") = Request("c1")
set rs = cmd.execute()
DispCur = rs("Currency")
CardName = rs("CardName")

sql = "select T0.LogNum " & _
	  "from r3_obscommon..tlog T0 " & _
	  "inner join r3_obscommon..tdoc T1 on T1.lognum = T0.Lognum " & _
	  "where Company = '" & Session("olkDB") & "' and T0.object in (" & ObjCode & ") " & sqlAdd & " and status in ('R', 'H') and T1.cardcode = N'" & saveHTMLDecode(Request("c1"), False) & "' " & _
	  "and T0.LogNum not in (select LogNum from OLKClientsDocControl where CardCode = N'" & saveHTMLDecode(Request("c1"), False) & "') order by T1.DocDate desc"
set rd = Server.CreateObject("ADODB.RecordSet")
rd.open sql, conn, 3, 1
rd.PageSize = 10
rd.CacheSize = 10

If Request("p") <> "" Then iCurPage = CInt(Request("p")) Else iCurPage = 1
iPageCount = rd.PageCount

If Not rd.Eof Then
	rd.AbsolutePage = iCurPage
	LogNum = ""
	For i = 1 to rd.PageSize
		If i > 1 Then LogNum = LogNum & ", "
		LogNum = LogNum & rd("LogNum")
		rd.movenext
		If rd.Eof Then Exit For
	Next
	
	sql = "select T0.LogNum, T1.CardCode, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRD', 'CardName', T1.CardCode, T2.CardName) CardName, Object, DocDate, Comments, " & _
		  "OLKCommon.dbo.DBOLKDocTotal" & Session("ID") & "(T0.LogNum) Total, T0.Status " & _
		  "from r3_obscommon..tlog T0 " & _
		  "inner join r3_obscommon..tdoc T1 on T1.lognum = T0.Lognum " & _
		  "inner join ocrd T2 on T2.cardcode = T1.cardcode collate database_default " & _
		  "where T0.LogNum in (" & LogNum & ") order by Convert(int,T1.DocDate) desc"
	set rd = conn.execute(sql)
End If

%><div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
    <tr>
      <td bgcolor="#9BC4FF">
      <table border="0" cellpadding="0" bordercolor="#111111" width="100%" id="AutoNumber1">
        <tr>
          <td width="100%">
          <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><%=getnewdocgoLngStr("LtxtPendList")%> 
          </font></b></td>
        </tr>
        <tr>
          <td width="100%" bgcolor="#3385FF">
			<table border="0" width="100%" id="table1" cellspacing="0" cellpadding="0">
				<tr>
					<td width="12"><a href="operaciones.asp?cmd=datos&card=<%=CleanItem(Request("c1"))%>"><img border="0" src="images/<%=Session("rtl")%>flechaselec.gif" align="left"></a></td>
					<td><font face="Verdana" size="1" color="#000000">&nbsp;<%=Request("C1")%> - <%=CardName%></font></td>
				</tr>
			</table>
			</td>
        </tr>
        <tr>
          <td width="100%">
          <table border="0" cellpadding="0"  bordercolor="#111111" width="100%" id="AutoNumber2">
            <tr>
              <td width="100%" colspan="3"><b><font size="1" face="Verdana">
              <p align="center"><%=Replace(getnewdocgoLngStr("LtxtPendingDocs"), "{0}", DocCount)%></font></b></td>
            </tr>
            <tr>
              <td width="44%"><b><font size="1" face="Verdana"></font></b></td>
              <td width="15%" bgcolor="#66A4FF"><b><font size="1" face="Verdana"><%=getnewdocgoLngStr("DtxtTotal")%>:</font></b></td>
              <td width="47%" bgcolor="#66A4FF">
              <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>" dir="ltr"><b><font size="1" face="Verdana"><nobr><%=DispCur%>&nbsp;<%=FormatNumber(DocTotal,myApp.SumDec)%></nobr></font></b></td>
            </tr>
            <% If myApp.EnableOQUT Then %>
            <tr>
              <td width="100%" colspan="3">
              <p align="center">
              <input type="button" value="<%=Replace(getnewdocgoLngStr("LtxtNewDoc"), "{0}", txtQuote)%>" onclick="javascript:window.location.href='newdocgonow.asp?c1=<%=Replace(CleanItem(request("c1")), "'", "\'")%>&ObjCode=23'">
              </td>
            </tr>
            <% End If %>
            <% If myApp.EnableORDR Then %>
            <tr>
              <td width="100%" colspan="3">
              <p align="center">
              <input type="button" value="<%=Replace(getnewdocgoLngStr("LtxtNewDoc"), "{0}", txtOrdr)%>" onclick="javascript:window.location.href='newdocgonow.asp?c1=<%=Replace(CleanItem(request("c1")), "'", "\'")%>&ObjCode=17'">
              </td>
            </tr>
            <% End If %>
            <% If myApp.EnableOINV Then %>
            <tr>
              <td width="100%" colspan="3">
              <p align="center">
              <input type="button" value="<%=Replace(getnewdocgoLngStr("LtxtNewDoc"), "{0}", txtInv)%>" onclick="javascript:window.location.href='newdocgonow.asp?c1=<%=Replace(CleanItem(request("c1")), "'", "\'")%>&ObjCode=13'">
              </td>
            </tr>
            <% End If %>
            <% If myApp.EnableOINVRes Then %>
            <tr>
              <td width="100%" colspan="3">
              <p align="center">
              <input type="button" value="<%=Replace(getnewdocgoLngStr("LtxtNewDoc"), "{0}", txtInvRes)%>" onclick="javascript:window.location.href='newdocgonow.asp?c1=<%=Replace(CleanItem(request("c1")), "'", "\'")%>&ObjCode=-13'">
              </td>
            </tr>
            <% End If %>
          </table>
          </td>
        </tr>
        <tr>
          <td width="100%"><hr color="#3385FF" size="1"></td>
        </tr>
        <% If Not rd.Eof Then
        do while not rd.eof
			Enable = True
			Select Case rd("Object")
		  		Case 17
		 			If Not myApp.EnableORDR Then Enable = False
		  		Case 23
		  			If Not myApp.EnableOQUT Then Enable = False
		  		Case 13
		  			If Not myApp.EnableOINV Then Enable = False
		  	End Select %>
        <tr>
          <td width="100%">
          <table border="0" cellpadding="0" cellspacing="1" width="100%" >
            <tr>
              <td width="9%" bgcolor="#66A4FF">
              <p align="center">
			    <a href="javascript:<% If Enable Then %>doGoDoc(<%=rd("Object")%>, '<%=rd("LogNum")%>', '<%=Replace(myHTMLEncode(rd("CardCode")), "'", "\'")%>', '<%=rd("status")%>');<% Else %>listPendAlert(<%=rd("Object")%>);<% End If %>">
			    <img border="0" src="images/<%=Session("rtl")%>flechaselec.gif" align="left"></a></td>
              <td width="20%" bgcolor="#66A4FF" style="width: 61%"><font size="1" face="Verdana"><%=RD("LogNum")%></font></td>
              <td width="30%" bgcolor="#66A4FF" align="right"><font size="1" face="Verdana"><%=FormatDate(RD("DocDate"), True)%></font></td>
            </tr>
            <tr>
              <td width="9%">
              <p align="center">
		    <a href="operaciones.asp?cmd=docdel&retval=<%=RD("LogNum")%>&c1=<%=CleanItem(Request("c1"))%>">
		    <img border="0" src="images/remove.gif"></a></td>
              <td width="20%" style="width: 61%"><font size="1" face="Verdana"><% 
              Select Case RD("Object") 
              	Case 17
	              	Response.Write txtOrdr
	            Case 23
	              	Response.Write txtQuote
	            Case 13
	            	Response.Write txtInv
              End Select %></font></td>
              <td width="30%"><p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>" dir="ltr"><nobr><font size="1" face="Verdana"><nobr><%=DispCur%>&nbsp;<%=FormatNumber(rd("Total"), myApp.SumDec)%></nobr></font></nobr></p></td>
    		</tr>
    		<% If rd("Comments") <> "" Then %>
    		<tr>
    			<td colspan="3"><font face="Verdana" size="1"><%=getnewdocgoLngStr("DtxtNote")%>: <%=rd("Comments")%></font>
    			</td>
    		</tr>
    		<% End If %>
          </table>
          </td>
        </tr>
	    <% rd.movenext
	    loop %>
        <tr>
          <td width="100%" colspan="4">
			<table border="0" cellpadding="0" cellspacing="1" width="100%" id="table1" dir="ltr">
				<tr>
					<td width="16">
					<% If iCurPage > 1 Then %><a href='newdocgo.asp?cmd=docgo&amp;c1=<%=Request("c1")%>&amp;p=<%=iCurPage-1%>'><img border="0" src="images/flecha_prev.gif" width="16" height="16"></a><% Else %>&nbsp;<% End If %></td>
					<td>
					<p align="center">
					<select name="cmbPage" size="1" onchange="javascript:window.location.href='operaciones.asp?cmd=docgo&c1=<%=Request("c1")%>&p=' + this.value">
					<% For i = 1 to iPageCount %>
					<option value="<%=i%>" <% If i = iCurPage Then %>selected<% End If %>><%=i%></option>
					<% Next %>
					</select></td>
					<td width="16">
					<% If iCurPage < iPageCount Then %><a href='newdocgo.asp?cmd=docgo&amp;c1=<%=Request("c1")%>&amp;p=<%=iCurPage+1%>'><img border="0" src="images/flecha_next.gif" width="16" height="16"></a><% End If %></td>
				</tr>
			</table>
			</td>
        </tr>
        <% End If %>
        <tr>
          <td width="100%">&nbsp;</td>
        </tr>
      </table>
      </td>
    </tr>
    </table>
  </center>
</div>
<script language="javascript">
function listPendAlert(obj) 
{
	var objType;
	switch (obj) {
		case 15:
			objType = "<%=txtOdlns%>";
			break;
		case 17:
			objType = "<%=txtOrdrs%>";
			break;
		case 23:
			objType = "<%=txtQuotes%>";
			break;
		case 24:
			objType = "<%=txtRcts%>";
			break;
		case 48:
			objType = "<%=txtInvs%>/<%=txtRcts%>";
			break;
		case 13:
			objType = "<%=txtInvs%>";
			break;
		case 4:
			objType = "|D:txtItem|";
			break;
		case 2:
			objType= "<%=txtClient%>"
			break;
	}
	alert('<%=getnewdocgoLngStr("LtxtDisObj")%>'.replace('{0}', objType));
}


function confReopen()
{
	return confirm('<%=getnewdocgoLngStr("LtxtConfReOpen")%>')
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
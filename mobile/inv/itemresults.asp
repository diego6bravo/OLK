<% addLngPathStr = "inv/" %>
<!--#include file="lang/itemresults.asp" -->
<%

iPageSize = 10

If Request("Page") = "" Then
	sql = "select T0.ItemCode, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'ItemName', T1.ItemCode, T1.ItemName) as ItemName, T1.OnHand " & _
	"from oitw T0 " & _
	"inner join oitm T1 on T1.itemcode = T0.itemcode "
	
	
	If myApp.EnableCodeBarsQry and myApp.CodeBarsQryMethod = "I" Then
		sql = sql & "cross join (" & Replace(myApp.CodeBarsQry, "@CodeBars", "N'" & saveHTMLDecode(Request("string"), False) & "'") & ") tCodeBars "
	End If
	
	sql = sql & "where whscode = N'" & Session("bodega") & "' and (T1.ItemCode like N'%" & Request("string") & "%' or "
	
  	If Not myApp.EnableCodeBarsQry Then
	  	sql = sql & "T1.CodeBars like N'%" & saveHTMLDecode(Request("string"), False) & "%'"
	Else
		Select Case myApp.CodeBarsQryMethod
			Case "R"
			  	sql = sql & "T1.CodeBars = (" & Replace(myApp.CodeBarsQry, "@CodeBars", "N'" & saveHTMLDecode(Request("string"), False) & "'") & ") "
			Case "I"
				sql = sql & "T1.CodeBars = tCodeBars.CodeBars "
		End Select
	  	
	End If
	
	If myApp.EnableSearchItmSupp Then
		sql = sql & " or T1.SuppCatNum = N'" & saveHTMLDecode(Request("string"), False) & "'"
	End If
	
	sql = sql & ")"

	Session("sqlstmt") = sql 
	rs.open sql, conn, 3, 1
Else
	sql = Session("sqlstmt")
	rs.open sql, conn, 3, 1
	iPageCurrent = CInt(Request("page"))
End If
if rs.recordcount = 1 then response.redirect "operaciones.asp?cmd=recountActiveItem&item=" & rs("ItemCode")

RS.PageSize = iPageSize
RS.CacheSize = iPageSize
iPageCount = RS.PageCount
If iPageCurrent > iPageCount Then iPageCurrent = iPageCount
If iPageCurrent < 1 Then iPageCurrent = 1

%><div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
    <!-- fwtable fwsrc="Z:\topmanage\logos\originales\pocket_art.png" fwbase="pocket_artpieza1.gif" fwstyle="FrontPage" fwdocid = "742308039" fwnested=""0" -->
    <tr>
      <td bgcolor="#9BC4FF">
      <table border="0" cellpadding="0" bordercolor="#111111" width="100%" id="AutoNumber1">
        <tr>
          <td width="100%">
          <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><%=getitemresultsLngStr("LtxtInvRecount")%>&nbsp; 
          </font></b></td>
        </tr>
        <% if rs.recordcount > 0 then
        RS.AbsolutePage = iPageCurrent %>
        <tr>
          <td width="100%" bgcolor="#75ACFF">
          <p align="center"><b><font size="1" face="Verdana"><%=getitemresultsLngStr("LtxtSelItmNote")%></font></b></td>
        </tr>
        <tr>
          <td width="100%">
          <table border="0" cellpadding="0" cellspacing="1"  bordercolor="#111111" width="100%" id="AutoNumber2">
             
            <tr>
              <td width="33%" align="center" bgcolor="#75ACFF" colspan="2"><b>
              <font size="1" face="Verdana"><%=getitemresultsLngStr("DtxtCode")%></font></b></td>
              <td width="34%" align="center" bgcolor="#75ACFF"><b>
              <font size="1" face="Verdana"><%=getitemresultsLngStr("LtxtInventory")%></font></b></td>
            </tr>
            <% Alter = False
            for intRecord=1 to rs.PageSize %>
            <tr <% If Alter Then %>bgcolor="#c1daff"<% End If %>>
              <td colspan="3" valign="top"><font size="1" face="Verdana"><%=RS("ItemName")%></font></td>
            </tr>
            <tr <% If Alter Then %>bgcolor="#c1daff"<% End If %>>
              <td width="7%" valign="top"><a href="operaciones.asp?cmd=recountActiveItem&item=<%=Replace(Replace(Replace(rs("ItemCode"),"#","%23"),"&","%26"),"""","%22")%>"><img border="0" src="images/<%=Session("rtl")%>flechaselec.gif"></a></td>
              <td width="26%" valign="top"><font size="1" face="Verdana"><%=RS("ItemCode")%></font></td>
              <td width="34%" valign="top"><font size="1" face="Verdana"><p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><%=RS("OnHand")%></font></td>
            </tr>
           <% Alter = Not Alter
           rs.movenext
        	if rs.EOF then exit for
			next 
			%>
          </table>
          </td>
        </tr>
        <tr>
          <td width="100%" colspan="2">
          <table border="0" cellpadding="0" cellspacing="1" bordercolor="#111111" width="100%" id="AutoNumber4" dir="ltr">
            <tr>
              <td width="8%" valign="top">
                <% If iPageCurrent > 1 Then %><a href="javascript:goPage(<%= iPageCurrent - 1 %>);"><img border="0" src="images/flecha_prev.gif"></a><% End If %></td>
              <td width="85%" align="center">
              <% If iPageCount > 0 Then %>
              <select name="pageSelection" size="1" onchange="javascript:goPage(this.value);" style="font-family: Verdana; font-size: 10px; border: 1px solid #5197FF; background-color: #9BC4FF">
              	<% For I = 1 to iPageCount %>
              	<option <% If I = iPageCurrent Then %>selected<% End If %> value="<%=i%>"><%=i%></option>
              	<% Next %>
              </select>
              <% End If %>
				</td>
              <td width="7%" valign="top">
              <% If iPageCurrent < iPageCount Then %><a href="javascript:goPage(<%= iPageCurrent + 1 %>);"><img border="0" src="images/flecha_next.gif"></a><% End If %></td>
            </tr>
          </table>
          </td>
        </tr>
        <% end if
        if rs.recordcount = 0 then %>
        <tr>
          <td width="100%">
			<p align="center"><font face="Verdana" size="1"><%=getitemresultsLngStr("DtxtNoData")%></font></td>
        </tr>
        <% end if %>
      </table>
      </td>
    </tr>
    </table>
  </center>
</div>
<form name="frmPage" method="post" action="operaciones.asp">
<% For each itm in Request.Form
If itm <> "Page" Then %>
<input type="hidden" name="<%=itm%>" value="<%=Request(itm)%>">
<% End If
Next %>
<input type="hidden" name="Page" value="">
</form>
<script language="javascript">
function goPage(p)
{
	document.frmPage.Page.value = p;
	document.frmPage.submit();
}
</script>
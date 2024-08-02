<% addLngPathStr = "cart/" %>
<!--#include file="lang/cart_cp.asp" -->
<%
set rx = Server.CreateObject("ADODB.RecordSet")
sql = "select T0.RowQuery, IsNull(T1.AlterRowName, T0.RowName) RowName, Align " & _
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
	sql = sql & "(" & rx("RowQuery") & ") As N'" & Replace(rx("RowName"), "'", "''") & "{S}" & rx("Align") & "'"
rx.movenext
loop
sql = QueryFunctions(sql)
set rx = conn.execute(sql)
End If %>
<center>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111" id="table9">
    <!-- fwtable fwsrc="Z:\topmanage\logos\originales\pocket_art.png" fwbase="pocket_artpieza1.gif" fwstyle="FrontPage" fwdocid = "742308039" fwnested=""0" -->
    <tr>
      <td bgcolor="#9BC4FF">
      <form method="POST" action="cart/cartupdate2.asp" name="frmCart">
        <table border="0" cellpadding="0" bordercolor="#111111" width="100%" id="table10">
        <tr>
          <td width="100%">
          <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><%=getcart_cpLngStr("LtxtShopCart")%>&nbsp;-&nbsp;<%=txtBasketMinRep%></font></b></td>
        </tr>
        <tr>
          <td width="100%" style="border-bottom-style: solid; border-bottom-width: 1px">
          <table border="0" cellpadding="0" bordercolor="#111111" width="100%" id="table11">
            <tr>
              <td width="100%">
          <table border="0" cellpadding="0" cellspacing="1"  bordercolor="#111111" width="100%" id="table12" style="font-family: Verdana; font-size: 10px">
          	<% If Not rx.eof Then
          	For each fld in rx.Fields
          	arrValues = Split(fld.Name, "{S}")
          	varName = arrValues(0)
          	varAlign = ""
          	Select Case arrValues(1)
          		Case "L"
          			varAlign = "left"
          		Case "C"
          			varAlign = "center"
          		Case "R"
          			varAlign = "right"
          	End Select %>
            <tr>
              <td width="28%" height="9" bgcolor="#95BFFF">
              <%=varName%></td>
              <td width="70%" height="9" bgcolor="#75ACFF" align="<%=varAlign%>" dir="ltr">
              <% If Not IsNull(fld) Then %><%=fld%><% End If %></td>
            	</tr>
            <% Next
            End If %>
            </table>
              </td>
            </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td width="100%" style="font-size: 5px; border-left-width: 1px; border-right-width: 1px; border-top: 1px solid #FF9933; border-bottom-width: 1px">&nbsp;</td>
        </tr>
        <tr>
        <td>
          <div align="center">
            <center>
            <table border="0" cellpadding="0" cellspacing="1" style="border-collapse: collapse" bordercolor="#111111" width="95" id="table13">
              <tr>
                <td>
                <p align="center"><a href="operaciones.asp?cmd=cart"><img border="0" src="images/ok_icon.gif"></a></td>
              </tr>
            </table>
            </center>
          </div>
        </td></TR>
    </table>
    </td></tr></table>
  </center>
</div>
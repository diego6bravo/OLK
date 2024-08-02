<!--#include file="chkLogin.asp" -->
<!--#include file="myHTMLEncode.asp" -->
<!--#include file="lang/selectClientes.asp" -->
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

          
	BannerID = Request("BannerID")
	
	If Request("hdnCodClientes") = "" Then
		strVerfy = "'N'"
	Else
		strVerfy = "Case When GroupCode in (" & Request("hdnCodClientes") & ")  Then 'Y' Else 'N' End "
	End If
	
	strSQL = "select GroupCode, GroupName, " & _
			strVerfy & " Verfy from OCRG " & _
			"where GroupType = 'C' "
	
	set rstClientes = Conn.Execute(strSQL)
%>
<html <% if session("rtl") <> "" then %>dir="rtl" <% end if %>>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="style/style_pop.css" rel="stylesheet" type="text/css">

<title><%=getselectClientesLngStr("LttlSelGrp")%></title>
<%
	'chkCodCliente contiene los índices de los clientes que serán utilizado para hacer el insert.
	'Print Request("Ejecutar")
	If Request("Ejecutar") = "S" Then
		cant = Request.Form("chkCodCliente").Count
		'Print cant
		If cant > 0 Then
			Redim arrCodCliente(cant-1)
			j = 0
			For i = 1 TO cant
				arrCodCliente(j) = Request.Form("chkCodCliente")(i)
				'Print arrCodCliente(j)&"<br>"
				j = j + 1
			Next
			strCodClientes = Join(arrCodCliente,", ")
		Else
			strCodClientes = "-1"	' Significa que no fue seleccionado ninguno.
		End If
%>
<script language="javascript">
		//Actualiza el control hidden que contiene todos los códigos de los clientes asociados.
		opener.pasarToCode('<%=strCodClientes%>', 'clientes')
		window.close()
</script>
<%
	End If
%>
<script type="text/javascript" src="general.js"></script>
<script language="javascript">
function setTblSet()
{
	if (browserDetect() == 'msie')
	{
		tblSave.style.top = document.body.offsetHeight-31+document.body.scrollTop;
	}
	else if (browserDetect() == 'opera')
	{
		tblSave.style.top = document.body.offsetHeight-27+document.body.scrollTop;
	}
	else //firefox & others
	{
		tblSave.style.top = window.innerHeight-27+document.body.scrollTop;
	}
}
</script>
</head>

<body topmargin="0" leftmargin="0" onbeforeunload="opener.clearWin();" onload="setTblSet();" onscroll="setTblSet();">

<table border="0" cellpadding="0" bordercolor="#111111" width="100%">
	<tr>
		<td class="popupTtl"><%=getselectClientesLngStr("LttlSelGrp")%>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" bordercolor="#111111" width="100%">
			<form method="POST" action="selectClientes.asp" name="frmClientes">
				<% Do while Not rstClientes.EOF %>
				<tr>
					<td class="popupOptValue">
					<input type="checkbox" name="chkCodCliente" class="noborder" id="chkCodCliente<%=rstClientes("GroupCode")%>" value="<%=rstClientes("GroupCode")%>" <%if rstclientes("Verfy") = "Y" then%> checked<%end if%>>
					<label for="chkCodCliente<%=rstClientes("GroupCode")%>"><%=Server.HTMLEncode(rstClientes("GroupName"))%></label>
					</td>
				</tr>
				<% rstClientes.MoveNext 
				Loop %>
				<tr height="27">
					<td>&nbsp;</td>
				</tr>
				<input type="hidden" value="S" name="Ejecutar">
			</form>
		</table>
		</td>
	</tr>
</table>
<table cellpadding="0" border="0" width="100%" style="position: absolute;" id="tblSave" bgcolor="#FFFFFF">
	<tr>
		<td style="width: 75px"><input type="button" value="<%=getselectClientesLngStr("DtxtSave")%>" name="cmdGuardar" class="OlkBtn" onclick="javascript:document.frmClientes.submit();"></td>
		<td>
			<hr color="#0D85C6" size="1"></td>
		<td style="width: 75px"><input type="button" value="<%=getselectClientesLngStr("DtxtClose")%>" name="cmdCerrar" onclick="javascript:window.close()" class="OlkBtn"></td>
	</tr>
</table>

</body>

</html>

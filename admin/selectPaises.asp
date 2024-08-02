<!--#include file="chkLogin.asp" -->
<!--#include file="myHTMLEncode.asp" -->
<!--#include file="lang/selectPaises.asp" -->
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

	BannerID = Request("BannerID")
	
	If Request("hdnCodPaises") = "" Then
		strVerfy = "'N'"
	Else
		strVerfy = "Case When Code in (" & Request("hdnCodPaises") & ") Then 'Y' Else 'N' End "
	End If
	
	strSQL = "select Code, Name, " & strVerfy & " Verfy " & _
				"from OCRY order by Name asc"
			 
	
	set rstPaises = Conn.Execute(strSQL)
	
%>
<html <% If Session("rtl") <> "" Then %>dir="rtl"<% End If %>>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="style/style_pop.css" rel="stylesheet" type="text/css">
<title><%=getselectPaisesLngStr("LttlSelCnt")%></title>

<%
	'chkCodPais contiene los índices de los Paises que serán utilizado para hacer el insert.
	'Print Request("Ejecutar")
	If Request("Ejecutar") = "S" Then
		cant = Request.Form("chkCodPais").Count
		'Print cant
		If cant > 0 Then
			Redim arrCodPaises(cant-1)
			j = 0
			For i = 1 TO cant
				arrCodPaises(j) = Request.Form("chkCodPais")(i)
				'Print arrCodPaises(j)&"<br>"
				j = j + 1
			Next
			strCodPaises = Join(arrCodPaises,", ")
		Else
			strCodPaises = "-1"	' Significa que no fue seleccionado ninguno.
		End If
%>
<script language="javascript">
//Actualiza el control hidden que contiene todos los códigos de los Paises asociados.
opener.pasarToCode("<%=strCodPaises%>", "paises");
window.close();
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
	
<table border="0" cellpadding="0"  bordercolor="#111111" width="100%">
	<tr>
		<td class="popupTtl"><%=getselectPaisesLngStr("LttlSelCnt")%></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0"  bordercolor="#111111" width="100%">
			<form method="POST" action="selectPaises.asp?BannerId=<%=BannerID%>" name="frmPaises">
<%
	Do while Not rstPaises.EOF
%>				
			<tr>
				<td class="popupOptValue">					
					<input type="checkbox" class="noborder" name="chkCodPais" id="chkCodPais<%=rstPaises("Code")%>" value="'<%=rstPaises("Code")%>'" <%If rstPaises("Verfy") = "Y" Then%> checked<%End If%>>
					<label for="chkCodPais<%=rstPaises("Code")%>"><%=Server.HTMLEncode(rstPaises("Name"))%></label>
				</td>
			</tr>
<%

		rstPaises.MoveNext 
	Loop
%>
	<tr height="27">
		<td>&nbsp;</td>
	</tr>
			<input type="hidden" value="S" name="Ejecutar">
			</form>
		</table>
		</td>
	</tr>
	</table>
	
<table cellpadding="0" border="0" width="98%" style="position: absolute;" id="tblSave" bgcolor="#FFFFFF">
	<tr>
		<td style="width: 75px"><input type="button" value="<%=getselectPaisesLngStr("DtxtSave")%>" name="cmdGuardar" class="OlkBtn" onclick="javascript:document.frmPaises.submit();"></td>
		<td>
			<hr color="#0D85C6" size="1"></td>
		<td style="width: 75px"><input type="button" value="<%=getselectPaisesLngStr("DtxtClose")%>" name="cmdCerrar" onclick="javascript:window.close()" class="OlkBtn"></td>
	</tr>
</table>
	
</body>

</html>
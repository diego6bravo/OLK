<!--#include file="chkLogin.asp" -->
<!--#include file="myHTMLEncode.asp" -->
<!--#include file="lang/selectSecciones.asp" -->
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

	BannerID = Request("BannerID")
	
	If Request("hdnCodSecciones") = "" Then
		strVerfy = "'N'"
	Else
		strVerfy = " Case When T0.SecType + Convert(nvarchar(20),T0.SecID) in (" & Request("hdnCodSecciones") & ") Then 'Y' Else 'N' End "
	End If
	
	strSQL = "select T0.SecID, T1.SecName, T0.SecType, " & strVerfy & " Verfy " & _
				"from OLKSections T0 " & _
				"inner join OLKCommon..OLKSectionsDesc T1 on T1.SecID = T0.SecID and T1.LanID = " & Session("LanID") & " " & _
				"where T0.SecID >= 0 and T0.Status <> 'D' " & _
				"Order by T0.SecOrder  "

	set rstSecciones = Conn.Execute(strSQL)
	
%>
<html <% If Session("rtl") <> "" Then %>dir="rtl"<% End If %>>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="style/style_pop.css" rel="stylesheet" type="text/css">
<title><%=getselectSeccionesLngStr("LttlSelSec")%></title>
<%
	'chkCodSeccion contiene los índices de los Secciones que serán utilizado para hacer el insert.
	'Print Request("Ejecutar")
	If Request("Ejecutar") = "S" Then
		cant = Request.Form("chkCodSeccion").Count
		'Print cant
		If cant > 0 Then
			'Redim arrCodSecciones(cant-1)
			'Redim arrSecTipos(cant-1)
			'j = 0
			'For i = 1 TO cant
			'	arrTemp = Split(Request.Form("chkCodSeccion")(i), ",")
			'	arrCodSecciones(j) = arrTemp(0)
			'	arrSecTipos(j) = arrTemp(1)
			'	j = j + 1
			'Next
			'strCodSecciones = Join(arrCodSecciones,",")
			'strSecTipos = Join(arrSecTipos,",")
			'Print strCodSecciones&"<br>"
			'Print strSecTipos&"<br>"
			'Response.End
			strCodSecciones = Request("chkCodSeccion")
		Else
			strCodSecciones = "-1"	' Significa que no fue seleccionado ninguno.
		End If
%>
<script language="javascript">
//Actualiza el control hidden que contiene todos los códigos de los Secciones asociados.
opener.pasarToCode("<%=strCodSecciones%>", "secciones", "<%=strSecTipos%>");
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

<table border="0" cellpadding="0" width="100%">
	<tr>
		<td class="popupTtl"><%=getselectSeccionesLngStr("LttlSelSec")%></td>
	</tr>
	<tr>
		<td height="71">
		<table border="0" cellpadding="0" width="100%">
			<form method="POST" action="selectSecciones.asp?BannerId=<%=BannerID%>" name="frmSecciones">
			<tr>
				<td class="popupOptDesc" style="width: 20px">					
					&nbsp;</td>
				<td class="popupOptDesc" align="center"><%=getselectSeccionesLngStr("DtxtTitle")%></td>
				<td class="popupOptDesc" align="center"><%=getselectSeccionesLngStr("DtxtType")%></td>
			</tr>			
			
<%
	Do while Not rstSecciones.EOF
%>				

			<tr class="popupOptValue">
				<td style="width: 20px">					
					<input type="checkbox" class="noborder" name="chkCodSeccion" id="chkCodSeccion<%=rstSecciones("SecType")%><%=rstSecciones("SecID")%>" value="'<%=rstSecciones("SecType")%><%=rstSecciones("SecID")%>'" <%If rstSecciones("Verfy") = "Y" Then%> checked<%End If%>>
				</td>
				<td><label for="chkCodSeccion<%=rstSecciones("SecType")%><%=rstSecciones("SecID")%>"><%=Server.HTMLEncode(rstSecciones("SecName"))%></label>
				</td>
				<td align=center>	
					<label for="chkCodSeccion<%=rstSecciones("SecID")%>,<%=rstSecciones("SecType")%>"><%
					If rstSecciones("SecType") = "U" Then
						Response.Write getselectSeccionesLngStr("DtxtUser")
					ElseIf rstSecciones("SecType") = "S" Then
						Response.Write getselectSeccionesLngStr("DtxtSystem")
					End If
					%></label>
				</td>

			</tr>
<%

		rstSecciones.MoveNext 
	Loop
%>
				<tr height="27">
					<td colspan="3">&nbsp;</td>
				</tr>
					<input type="hidden" value="S" name="Ejecutar">
			</form>
		</table>
		</td>
	</tr>
	</table>
<table cellpadding="0" border="0" width="100%" style="position: absolute;" id="tblSave" bgcolor="#FFFFFF">
	<tr>
		<td style="width: 75px"><input type="button" value="<%=getselectSeccionesLngStr("DtxtSave")%>" name="cmdGuardar" class="OlkBtn" onclick="javascript:document.frmSecciones.submit();"></td>
		<td>
			<hr color="#0D85C6" size="1"></td>
		<td style="width: 75px"><input type="button" value="<%=getselectSeccionesLngStr("DtxtClose")%>" name="cmdCerrar" onclick="javascript:window.close()" class="OlkBtn"></td>
	</tr>
</table>
	
</body>

</html>
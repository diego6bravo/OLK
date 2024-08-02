<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>Seleccione la imgen que desea subir</title>
</head>

<body>

<form method="POST" action="subirfile.asp" webbot-action="--WEBBOT-SELF--">
  <!--webbot bot="SaveResults" u-file="../../../_private/form_results.csv" s-format="TEXT/CSV" s-label-fields="TRUE" startspan --><input NAME="VTI-GROUP" TYPE="hidden" VALUE="0"><!--webbot bot="SaveResults" i-checksum="37496" endspan --><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="365" id="AutoNumber1">
  <tr>
    <td>
    <table border="0" cellpadding="0"  bordercolor="#111111" width="100%" id="AutoNumber2">
      <tr>
        <td width="100%" bgcolor="#D2E9FF"><b><font face="Verdana" size="1">&nbsp;Seleccione la 
        imgen que desea subir</font></b></td>
      </tr>
      <tr>
        <td width="100%">
        <table border="0" cellpadding="0" cellspacing="1"  bordercolor="#111111" width="100%" id="AutoNumber3">
          <tr>
            <td width="33%">
            <input type="file" name="T1" size="32" style="font-family: Verdana; font-size: 10px; color: #3366CC; background-color: #E6F3FF"></td>
            <td width="34%" bgcolor="#F0F8FF">&nbsp;<input type="button" value="Subir &gt;&gt;" name="B1" style="font-family: Verdana; font-size: 10px"></td>
          </tr>
        </table>
        </td>
      </tr>
      <tr>
        <td width="100%" bgcolor="#F0F8FF">
	<table border="0" cellspacing="0" width="365" style="font-family: Verdana; font-size: 10px; border-collapse:collapse" bordercolor="#111111" cellpadding="0">
	<tr><td width="124"><span lang="en-us"><b>Detalles del archivo</b></span></td>
	<td>
	<p align="right">
	&nbsp;</td>	
	</tr>
	<tr>
	<td width="124"><b><span lang="en-us">Nombre</span> :</b></td>
	<td><asp:label id="fname" text="" runat="server" Font-Names="Verdana" Font-Size="XX-Small" /></td></tr>
	<tr>
	<td width="124"><b><span lang="en-us">Codificaci</span><span lang="es-pa">n</span> 
    :</b></td>
	<td><asp:label id="fenc" runat="server" Font-Names="Verdana" Font-Size="XX-Small" /></td></tr>
	<tr>
	<td width="124"><b><span lang="es-pa">Tamao</span> :(<span lang="es-pa">en</span> 
    bytes)</b></td>
	<td><asp:label id="fsize" runat="server" Font-Names="Verdana" Font-Size="XX-Small" /></td></tr>
	</table>
	    </td>
      </tr>
      <tr>
        <td width="100%">
<div align="center">
	<table border="0" cellpadding="0" cellspacing="0" width="365" id="table1">
		<tr>
			<td bgcolor="#F0F8FF">
			<p align="right" dir="ltr">
			<asp:label id="errmsg" runat="server" Font-Names="Verdana" Font-Size="XX-Small" ForeColor="#9D0400" />
			<input type="button" value="Cancelar" style="font-family: Verdana; font-size: 10px" runat="server" onclick="window.close()" ><input type="button" value="Aceptar" OnServerClick="AcceptFile" runat="server" style="font-family: Verdana; font-size: 10px" ></td>
		</tr>
	</table>
</div>
        </td>
      </tr>
    </table>
    </td>
  </tr>
</table>

</form>

</body>

</html>
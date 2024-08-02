<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>Subir Imagen</title>
<script LANGUAGE="VBScript">
<!-- Option Explicit
dim validation
dim header
header = "OLK"
Function Form1_OnSubmit
validation = True
if document.form1.File1.value = "" then
MsgBox "Tienes que escoger una imag�n",8, Header
validation = False
End If
If validation = True Then
Form1_OnSubmit = True
Else
Form1_OnSubmit = False
End If
End Function
-->
</script>
</Head>
<body topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0">
<FORM method="post" name="form1" encType="multipart/form-data" action="ToDatabase.asp">
    <table border="0" cellpadding="0"  bordercolor="#111111" width="100%" id="AutoNumber2">
 <tr>
    <td>
    <table border="0" cellpadding="0"  bordercolor="#111111" width="100%" id="AutoNumber2">
      <tr>
        <td width="100%" bgcolor="#D2E9FF"><font face="Verdana" size="1"><b>&nbsp;Seleccione la 
        imagen que desea subir:<br>
		</b>*.gif *.jpg *.png</font></td>
      </tr>
      <tr>
        <td width="100%" height="30" valign="middle">
            	<INPUT type="File" name="File1" style="font-family: Verdana; font-size: 10px" size="41" style="border:1px solid #68A6C0; font-family: Verdana; font-size: 10px; background-color:#E5F1FF">
			<INPUT type="submit" value="Subir" style="border:1px solid #68A6C0; font-family: Verdana; font-size: 10px; background-color:#E5F1FF"></td>
      </tr>
      </table>
    </td>
  </tr>
</table>
</FORM>
</body>
</html>
<html>
<head>
<title>Subir Imagen</title>
<script language="C#" runat="server">
void Upload(object Source, EventArgs e)
{

   if (myFile.PostedFile != null)
   {
   	try
   	{
      string strFileNamePath;
      string strFileNameOnly;
      string strTargetPath;

      strFileNamePath = myFile.PostedFile.FileName;
      strFileNameOnly = System.IO.Path.GetFileName(strFileNamePath);
      strTargetPath = Request["path"].ToString().Replace("{BS}","\\");

      myFile.PostedFile.SaveAs(strTargetPath + strFileNameOnly);
      Response.Redirect("acceptFile.asp?filename=" + strFileNameOnly);
	}
	catch (Exception ex)
	{
		string ErrMsg = "<script language=\"javascript\">" +
						"alert('Hubo un error al actualizar el archivo: " + ex.Message.Trim().Replace("'", "\\'") + "');<" +
						"/script>";
		Page.RegisterClientScriptBlock("Alert1", ErrMsg);
	}
   }
}

</script>

</head>
<body bgcolor="#F0F8FF" topmargin="0" leftmargin="0" onbeforeunload="opener.clearWin();">
<table border="0" cellpadding="0" width="100%" id="table1">
	<tr>
		<td bgcolor="#D6EAFE"><b><font face="Verdana" size="1">Subir Imagen</font></b></td>
	</tr>
	<tr>
		<td bgcolor="#EAF5FF">
		<form enctype="multipart/form-data" runat="server">
		   <font face="Verdana" size="1">Archivo:</font> 
			<input id="myFile" type="file" runat="server" style="font-family: Verdana; font-size: 10px" name="F1" size="20">
		   <input type=button value="Subir" OnServerClick="Upload" runat="server" style="font-family: Verdana; font-size: 10px">
		</form>
		</td>
	</tr>
</table>

</body>

</html>
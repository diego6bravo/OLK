<%@ Page Language="C#" AutoEventWireup="true" Debug="True"  %>
<%@ Import Namespace="System.Xml" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" id="myHTML" runat="server">

<head runat="server">
<title id="myTitle" runat="server"></title>
<link rel="stylesheet" href="<%=Request["style"]%>"/>
</head>
<!--#include file="lang/fileupload.aspx" -->
<body topmargin="0" leftmargin="0" onbeforeunload="opener.clearWin();">
<script runat="server">
protected void Page_Load(object sender, EventArgs e)
{
    loadLanguage();
	if (!IsPostBack)
	{
        Button1.Text = getLangVal("LtxtSend");
        lblUpload.Text = getLangVal("LtxtUpload");
        myTitle.Text = getLangVal("LtxtSendFile");
        btnClose.Text = getLangVal("DtxtClose");
        if (getLng() == "he") myHTML.Attributes.Add("dir", "rtl");
        
        if (Request["Source"] != null)
        {
        	if ((string)Request["Source"] == "Admin")
    		{
    			Button1.CssClass = "OlkBtn";
    			btnClose.CssClass = "OlkBtn";
    			tdTtl.Attributes.Add("class", "popupTtl");
    			tdFile.Attributes.Add("class", "popupOptValue");
        	}
        	else
        	{
				tdTtl.Attributes.Add("class", "GeneralTblBold2");
				tdFile.Attributes.Add("class", "GeneralTbl");
        	}
        }
    	else
    	{
			tdTtl.Attributes.Add("class", "GeneralTblBold2");
			tdFile.Attributes.Add("class", "GeneralTbl");
    	}

	}
}
string strFileNameOnly;
protected void Button1_Click(object sender, EventArgs e)
{
    if (FileUpload1.HasFile)
    {
        string fileExt = System.IO.Path.GetExtension(FileUpload1.FileName);
        fileExt = fileExt.ToLower();
        if (fileExt.Equals(".jpg") || fileExt.Equals(".gif") || fileExt.Equals(".png") || fileExt.Equals(".bmp"))
        {
            float lngTamano = FileUpload1.FileBytes.Length / 1024.0F;
            if (lngTamano > 1024)    //1 mega
            {
            	lblMensaje.Text = getLangVal("LtxtMinSize");
               	lblMensaje.ForeColor = System.Drawing.Color.Red;
            }
            else
            {
            	string strpath = "";
                try
                {
	                string connStr = System.Configuration.ConfigurationManager.ConnectionStrings["ConnStr"].ConnectionString;
	                SqlConnection sqlCn = new SqlConnection(connStr);
	                SqlCommand sqlCm = new SqlCommand("OLKGetImagePath", sqlCn);
		            sqlCm.CommandType = CommandType.StoredProcedure;
	                sqlCm.Parameters.Add("@ID", SqlDbType.Int).Value = Convert.ToInt32(Request["ID"]);
	                sqlCn.Open();
	                strpath = sqlCm.ExecuteScalar().ToString();
	                sqlCn.Close();
	                
	                strFileNameOnly = FileUpload1.FileName;
	                strpath = strpath.ToString();
	                strpath = strpath.Replace("{BS}","\\");
	                lblMensaje.Text = strpath;
	                FileUpload1.SaveAs(strpath + strFileNameOnly);
	                lblMensaje.Text = "File name: " +
	                FileUpload1.PostedFile.FileName + "<br>" +
	                FileUpload1.PostedFile.ContentLength + " kb<br>" +
	                "Content type: " +
	                FileUpload1.PostedFile.ContentType;
	                //Response.Redirect("acceptFile.asp?filename=" + Server.URLEncode(strFileNameOnly));
	                divUpload.Visible = false;
	                phClose.Visible = true;
                }
                catch (Exception ex)
                {
                    lblMensaje.Text = string.Format("Error: {0}<br/>ID: {1}<br/>Path: {2}", ex.Message, Request["ID"], strpath);
                }
            }
        }
        else
        {
        	lblMensaje.Text = getLangVal("LtxtFileTypes");
            lblMensaje.ForeColor = System.Drawing.Color.Red;
        }
    }	
    else
    {
    	lblMensaje.Text = getLangVal("LtxtSelFile");
    }
}
</script>
<form id="form1" runat="server">
	<div id="divUpload" runat="server">
		<table border="0" cellpadding="0" width="100%" id="table1">
			<tr>
				<td id="tdTtl" runat="server">
				<asp:label id="lblUpload" runat="server">Subir Imagen</asp:label>
				</td>
			</tr>
			<tr>
				<td id="tdFile" runat="server" width="100%">
				<asp:FileUpload ID="FileUpload1" runat="server" style="width: 100%; " />
				<br />
				<asp:Label ID="lblMensaje" runat="server"></asp:Label>
				</td>
			</tr>
			<tr>
				<td>
				<table cellpadding="0" border="0" width="100%">
					<tr>
						<td style="width: 1px; "><asp:Button ID="Button1" runat="server" OnClick="Button1_Click" CssClass="OlkBtn" Text="Subir" /></td>
						<td><hr size="1"/></td>
						<td style="width: 1px; "><asp:Button ID="btnClose" runat="server" CssClass="OlkBtn" Text="Cerrar" OnClientClick="window.close();" /></td>
					</tr>
				</table>
				</td>
			</tr>
		</table>
	</div>
	<asp:PlaceHolder ID="phClose" runat="server" Visible="false">
	<script language="javascript">
	opener.changepic('<%=strFileNameOnly%>');
	window.close()
	</script>
	</asp:PlaceHolder>
</form>

</body>

</html>

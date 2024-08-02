<%@ Page Language="C#" Debug="true" %>
<%@ Import Namespace=System.Drawing %>
<%@ Import Namespace=System %>
<%@ Import Namespace=System.Web %>
<html>
<script runat="server">
void Page_Load(object sender, System.EventArgs e)
{
	System.Drawing.Image originalimg;
	string FileName = Request.QueryString["filename"];
	string imgPath = "";
    if (Request.QueryString["dbName"] == null)
    {
    	imgPath = "Imagenes/" +  Request.Cookies["OLKDB"].Value.Replace("%5F","_") + "/";
    }
    else
    {
    	imgPath = "Imagenes/" +  Request.QueryString["dbName"] + "/";
    }
	int MaxSize = 100;
	if (Request.QueryString["MaxSize"] != null) MaxSize = Convert.ToInt16(Request.QueryString["MaxSize"]);
	int w,h;
	
	try
	{
		originalimg = System.Drawing.Image.FromFile(Server.MapPath(imgPath + FileName));
	}
	catch
	{
		originalimg = System.Drawing.Image.FromFile(Server.MapPath(imgPath + "n_a.gif"));
	}
	w = originalimg.Width; h = originalimg.Height;
	if (Request["wide"] != null) if (Request["wide"] == "Y") if (w > h) MaxSize = MaxSize + Convert.ToInt32(Request["MaxDif"])-45;

	if (w < MaxSize && h < MaxSize && Request["isZoom"] == null) {}
	else if (w > h){ h = (h * MaxSize)/w; w = MaxSize; }
	else if (h > w){ w = (w * MaxSize)/h; h = MaxSize; }
	else if (h == w){w = MaxSize;h = MaxSize;}
	else if (Request["isZoom"] != null)
	{
		if ((string)Request["isZoom"] == "Y")
		{
		}
	}
	
	Bitmap bmPhoto = new Bitmap(w,h,System.Drawing.Imaging.PixelFormat.Format32bppRgb);
	bmPhoto.SetResolution(originalimg.HorizontalResolution,originalimg.VerticalResolution);

	Graphics grPhoto = Graphics.FromImage(bmPhoto);
	grPhoto.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
	
	grPhoto.DrawImage(originalimg,
		new Rectangle(0,0,w,h),
		new Rectangle(0,0,originalimg.Width,originalimg.Height),
		GraphicsUnit.Pixel);

	grPhoto.Dispose();

    Response.ContentType = "image/jpeg";
    bmPhoto.Save(Response.OutputStream, System.Drawing.Imaging.ImageFormat.Jpeg);
    
	originalimg.Dispose();
	bmPhoto.Dispose();
}
</script>
</html>
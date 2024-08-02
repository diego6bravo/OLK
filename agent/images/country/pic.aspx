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

	int MaxHeight = 25;
	if (Request.QueryString["MaxHeight"] != null) MaxHeight = Convert.ToInt16(Request.QueryString["MaxHeight"]);
	int w,h;
	
	try
	{
		originalimg = System.Drawing.Image.FromFile(Server.MapPath(FileName));
	}
	catch
	{
		originalimg = System.Drawing.Image.FromFile(Server.MapPath("blank.gif"));
	}
	w = originalimg.Width; 
	h = originalimg.Height;

	w = (w * MaxHeight)/h; h = MaxHeight;
	
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
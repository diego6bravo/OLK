' Author: Karthik Giddu
' Email: gidduk@vsnl.com
' Updated: 5-May-2002
' Language: VB.NET
' Framework Version: V1
' Create thumbnail images on the fly which will double as image conversion (bmp to jpg/gif etc) utility
<%@ Page Language="VB" Debug="true" %>
<%@ Import Namespace=System.Drawing %>
<%@ Import Namespace=System %>
<%@ Import Namespace=System.Web %>

<html>
<script language="VB" runat="server">

  Sub Page_Load(Sender As Object, E As EventArgs)
	
        Dim orginalimg, thumb As System.Drawing.Image
        Dim FileName As String
        Dim inp As New IntPtr()
        Dim width, height As Integer
        Dim rootpath As String
        Dim OLKImgPath As String
        Dim MaxSize as Integer
        
        OLKImgPath = "imagenes/"
        
        rootpath = Server.MapPath(OLKImgPath) ' Get Root Application Folder

        FileName = Server.MapPath(Request.QueryString("FileName")) ' Root Folder + FileName
        Try
            orginalimg = orginalimg.FromFile(FileName) ' Fetch User Filename
        Catch
            orginalimg = orginalimg.FromFile(rootpath & "error.gif") ' Fetch error.gif
        End Try

        ' Get width using QueryString.
        If Request.QueryString("width") = Nothing Then
            width = orginalimg.Width  ' Use Orginal Width. 
        ElseIf Request.QueryString("width") = 0 Then  ' Assign default width of 100.
            width = 100
        Else
            width = Request.QueryString("width") ' Use User Specified width.
        End If

        ' Get height using QueryString.
        If Request.QueryString("height") = Nothing Then
            height = orginalimg.Height ' Use Orginal Height.
        ElseIf Request.QueryString("height") = 0 Then ' Assign default height of 100.
            height = 100
        Else
            height = Request.QueryString("height") ' Use User Specified height.
        End If
        
        If Request.QueryString("MaxSize") <> "" Then MaxSize = Request.QueryString("MaxSize") Else MaxSize = 100
        
        If width < MaxSize and height < MaxSize Then
        ElseIf width > height Then
			height = (height*MaxSize)/width
			width = MaxSize
        ElseIf height > width Then
			width = (width*MaxSize)/Height
			height = MaxSize
		ElseIf height = width
			width = MaxSize
			height = MaxSize
        End If
        
        thumb = orginalimg.GetThumbnailImage(width, height, Nothing, inp)

        ' Sending Response JPEG type to the browser. 
        Response.ContentType = "image/jpeg"
        thumb.Save(Response.OutputStream, Imaging.ImageFormat.Jpeg)

        ' Disposing the objects.
        orginalimg.Dispose()
        thumb.Dispose()

  End Sub
</script>
</html>

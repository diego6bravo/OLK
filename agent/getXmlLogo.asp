<%
	imgPath = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL")
	imgPath = Replace(imgPath, "getXmlLogo.asp", Request("logoImg"))
%><?xml version="1.0" encoding="iso-8859-1"?>
<imagenes>
<imagen id="<%=imgPath%>"/>
</imagenes>

<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="RTE_configuration/RTE_skin_file.asp" -->
<!--#include file="language_files/RTE_language_file_inc_es.asp" -->
<%
'****************************************************************************************
'**  Copyright Notice    
'**
'**  Web Wiz Guide - Web Wiz Rich Text Editor
'**  http://www.richtexteditor.org
'**                                                              
'**  Copyright 2002-2005 Bruce Corkhill All Rights Reserved.                                
'**
'**  This program is free software; you can modify (at your own risk) any part of it 
'**  under the terms of the License that accompanies this software and use it both 
'**  privately and commercially.
'**
'**  All copyright notices must remain in tacked in the scripts and the 
'**  outputted HTML.
'**
'**  You may use parts of this program in your own private work, but you may NOT
'**  redistribute, repackage, or sell the whole or any part of this program even 
'**  if it is modified or reverse engineered in whole or in part without express 
'**  permission from the author.
'**
'**  You may not pass the whole or any part of this application off as your own work.
'**   
'**  All links to Web Wiz Guide and powered by logo's must remain unchanged and in place
'**  and must remain visible when the pages are viewed unless permission is first granted
'**  by the copyright holder.
'**
'**  This program is distributed in the hope that it will be useful,
'**  but WITHOUT ANY WARRANTY; without even the implied warranty of
'**  MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE OR ANY OTHER 
'**  WARRANTIES WHETHER EXPRESSED OR IMPLIED.
'**
'**  You should have received a copy of the License along with this program; 
'**  if not, write to:- Web Wiz Guide, PO Box 4982, Bournemouth, BH8 8XP, United Kingdom.
'**    
'**
'**  No official support is available for this program but you may post support questions at: -
'**  http://www.webwizguide.info/forum
'**
'**  Support questions are NOT answered by e-mail ever!
'**
'**  For correspondence or non support questions contact: -
'**  info@webwizguide.info
'**
'**  or at: -
'**
'**  Web Wiz Guide, PO Box 4982, Bournemouth, BH8 8XP, United Kingdom
'**
'****************************************************************************************

Response.AddHeader "pragma","cache"
Response.AddHeader "cache-control","public"
Response.CacheControl = "Public"

%>
<html>
<head>
<title>No Preview</title>
<!--#include file="RTE_configuration/browser_page_encoding_inc.asp" -->

<%
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
Response.Write(vbCrLf & "<!--//" & _
vbCrLf & "/* *******************************************************" & _
vbCrLf & "Application: Web Wiz Rich Text Editor" & _
vbCrLf & "Author: Bruce Corkhill" & _
vbCrLf & "Info: http://www.richtexteditor.org" & _
vbCrLf & "Available FREE: http://www.richtexteditor.org" & _
vbCrLf & "Copyright: Bruce Corkhill ©2001-2005. All rights reserved" & _
vbCrLf & "******************************************************* */" & _
vbCrLf & "//-->")
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
%>

<script language="JavaScript">

//function to upadte image properties
function imageProperties(oImage){

	if (document.getElementById('prevFile').width != 1 && document.getElementById('prevFile').height != 1){
		window.parent.document.getElementById('width').value = document.getElementById('prevFile').width
		window.parent.document.getElementById('height').value = document.getElementById('prevFile').height
	}
}
function clearWin()
{
}
</script>
<style type="text/css">
<!--
.text {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 11px;
	color: #000000;
}
html,body { 
	border: 0px; 
}
-->
</style>
</head>
<body bgcolor="#FFFFFF" leftmargin="2" topmargin="2" marginwidth="2" marginheight="2">
<img src="<% = strRTEImagePath %>clear_pixel.gif" id="prevFile" onError="alert('<% = strTxtErrorLoadingPreview %>')" onLoad="imageProperties(this)"><span class="text">TopManage de Panam&#225;, brind&#225;ndole a sus clientes lo ultimo en tecnolog&#237;a y desarrollo para el crecimiento de su negocio, ofrece un producto innovador y &#250;nico en su contenido y funcionalidad, el OLK. 

Su compa&#241;&#237;a tendr&#225; un sitio web exclusivo para comercio electr&#243;nico. Conquiste un nuevo mercado y maximice su negocio brindando nuevos servicios a sus clientes y agentes con cat&#225;logos, pedidos en l&#237;nea siempre actualizados y comunicados directamente con informaci&#243;n de inventarios, costos, cantidades disponibles y hasta fotograf&#237;as de sus productos; ofreci&#233;ndole el 100% de seguridad en la informaci&#243;n </span>
</body>
</html>

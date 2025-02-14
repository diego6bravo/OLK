<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="RTE_configuration/RTE_setup.asp" -->
<!--#include file="RTE_configuration/RTE_skin_file.asp" -->
<!--#include file="functions/RTE_functions_common.asp" -->
<!--#include file="language_files/RTE_language_file_inc_es.asp" -->
<!--#include file="functions/functions_upload.asp" -->
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


'Set the timeout of the page
Server.ScriptTimeout =  1000


'Set the response buffer to true as we maybe redirecting
Response.Buffer = True


'Declare variables
Dim strErrorMessage	'Holds the error emssage if the file is not uploaded
Dim lngErrorFileSize	'Holds the file size if the file is not saved because it is to large
Dim blnExtensionOK	'Set to false if the extension of the file is not allowed
Dim strImageName	'Holds the file name
Dim saryFileUploadTypes	'Holds the array of file to upload
Dim strTextAreaName



'Intiliase variables
blnExtensionOK = True
strTextAreaName = Request.QueryString("textArea")




'If this is a post back then upload the image
If Request.QueryString("PB") = "Y" Then
	
	'Get the image types to upload
	saryFileUploadTypes = Split(Trim(strImageTypes), ";")
	
	'Call upoload file function
	strImageName = fileUpload(strImageUploadPath, saryFileUploadTypes, intMaxImageSize, strUploadComponent, lngErrorFileSize, blnExtensionOK)

End If


%>
<html>
<head>
<meta name="copyright" content="Copyright (C) 2002-2005 Bruce Corkhill" />
<title>Image Upload</title>
<!--#include file="RTE_configuration/browser_page_encoding_inc.asp" -->

<%
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
Response.Write(vbCrLf & "<!--//" & _
vbCrLf & "/* *******************************************************" & _
vbCrLf & "Application: Web Wiz Rich Text Editor ver. " & strRTEversion & "" & _
vbCrLf & "Author: Bruce Corkhill" & _
vbCrLf & "Info: http://www.richtexteditor.org" & _
vbCrLf & "Available FREE: http://www.richtexteditor.org" & _
vbCrLf & "Copyright: Bruce Corkhill �2001-2005. All rights reserved" & _
vbCrLf & "******************************************************* */" & _
vbCrLf & "//-->")
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
%>

<script  language="JavaScript">
<%
'If the image has been saved then place it in the post
If lngErrorFileSize = 0 AND blnExtensionOK = True AND strImageName <> "" Then
	%>

	window.opener.document.getElementById('<% = strTextAreaName %>').focus();
	window.opener.document.getElementById('<% = strTextAreaName %>').value += '<img src="<% = strFullURLpathToRTEfiles & Replace(strImageUploadPath, "\", "/", 1, -1, 1)  & "/" & strImageName %>" />';
	window.close();
<%
End If
%>
</script>
<style type="text/css">
<!--
html, body {
  background: ButtonFace;
  color: ButtonText;
  font: font-family: Verdana, Arial, Helvetica, sans-serif;
  font-size: 12px;
  margin: 1px;
  padding: 3px;
}
legend {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 12px;
	color: #0000FF;
}
-->
</style>
<link href="RTE_configuration/default_style.css" rel="stylesheet" type="text/css" />
</head>
<body OnLoad="self.focus(); document.forms.frmImageUp.Submit.disabled=true;"><%

'If the user is allowed to upload then show them the form
If blnImageUpload Then	

	%>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="0">
 <form action="non_RTE_upload_images.asp?PB=Y&textArea=<% = Server.URLEncode(strTextAreaName) %>" method="post" enctype="multipart/form-data" name="frmImageUp" target="_self" id="frmImageUp" onSubmit="alert('<% = strTxtPleaseWaitWhileImageIsUploaded %>')">
  <tr> 
   <td width="100%"> 
    <fieldset>
    <legend><% = strTxtImageUpload %></legend>
    <table width="100%" border="0" cellspacing="0" cellpadding="1">
     <tr> 
      <td width="10%" align="right" class="text"><% = strTxtImage %>:</td>
      <td width="90%"><input name="file" type="file" size="35" onFocus="document.forms.frmImageUp.Submit.disabled=false;" onChange="document.forms.frmImageUp.Submit.disabled=false;" />
        </td>
     </tr>
     <tr align="center"> 
      <td colspan="2" class="text"><br /><% 
      	
      	'If the file upload has failed becuase of the wrong extension display an error message
	If blnExtensionOK = False Then

		Response.Write("<span class=""error"">" & strTxtImageOfTheWrongFileType & ".<br />" & strTxtImagesMustBeOfTheType & ", "  &  Replace(strImageTypes, ";", ", ", 1, -1, 1) & "</span>")

	'Else if the file upload has failed becuase the size is to large display an error message
	ElseIf lngErrorFileSize <> 0 Then

		Response.Write("<span class=""error"">" & strTxtImageFileSizeToLarge & " " & lngErrorFileSize & "KB.<br />" & strTxtMaximumFileSizeMustBe & " " & intMaxImageSize & "KB</span>")
	
	'Else display a message of the image types that can be uploaded
	Else
      
      		Response.Write(strTxtImagesMustBeOfTheType & ", " & Replace(strImageTypes, ";", ", ", 1, -1, 1) & ", " & strTxtAndHaveMaximumFileSizeOf & " " & intMaxImageSize & "KB") 
      	
      	End If
      %></td>
     </tr>
    </table>
    </fieldset></td>
  </tr>
  <tr align="right"> 
   <td> <input type="submit" name="Submit" value="     <% = strTxtOK %>     "> &nbsp; <input type="button" name="cancel" value=" <% = strTxtCancel %> " onClick="window.close()"></td>
  </tr>
 </form>
</table><%

End If

%>
</body>
</html>

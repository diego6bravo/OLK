<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" >
<title>OLK Login Demo</title>
</head>

<body>

<p align="center"><font size="6" face="Verdana"><br>
<br>
<br>
<br>
OLK Login Demo</font></p>
<p align="center">
<% 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' If you want to change the look and feel of the component,					''
'' open the file "login.asp" in the login coponent folder						''
'' and change the style of anything you want									''
'' Alert: Don't change any server side code, you are only allowed 				''
'' to change fonts, colors and sizes of the control components					''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

LoginPath = "OLKLogin/" 'Enter the root where the login component files are located
Dim myLogin
set myLogin = new OLKLogin
myLogin.Language = "en" 'Enter the login language, the following options are allowed: en, es, he, pt, fr, de, ru
myLogin.EnableSavePassword = True 'Set if you want to allow the user to save his password
myLogin.AllowChangeUserType = True 'Set if you want to allow the user to select the user type
myLogin.DefaultUserType = "C" 'Set the default user type (mandatory if AllowChangeUserType equals False), the following options are allowed, C = Client, V = Vendor/Agent
myLogin.AccessKey = "JVI6XJC9CH4EMJ9AF7DSK6HSX9597IAFBNSCNQQLYULF5H2VLX" 	'Go to the Administration, Access, External login, 
																			'add the source URL from where the login is going to show and add it. 
																			'Copy the new access key for that URL and Paste it here.
myLogin.Database = "ESHOPS" 'Set the database property where you created the access key
myLogin.TargetURL = "http://localhost/newDeploy/" 'Enter the target URL where OLK Clients/Agents is hosted %>
<!--#include file="OLKLogin/login.asp" --><!-- Include File Connects this page to the OLK Login Component -->
<% myLogin.GenerateLogin 'Execute the generate login process %>
</p>
</body>

</html>

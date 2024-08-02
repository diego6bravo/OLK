<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<!--#include file="authorizationClass.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

Dim myAut
set myAut = New clsAuthorization

strScriptName = LCase(Request.ServerVariables("SCRIPT_NAME"))
If InStr(strScriptName ,"/") > 0 Then 
	strScriptName = right(strScriptName, len(strScriptName) - InStrRev(strScriptName,"/")) 
End If 

strRootPath = Replace(LCase(Request.ServerVariables("URL")), strScriptName, "")


CheckLanguage

%>
<!--#include file="lcidReturn.inc"-->
<!--#include file="controls.inc"-->
<!-- #INCLUDE file="FCKeditor/fckeditor.asp" -->
<!-- #include file="chkLogin.asp" -->
<!--#include file="lang.asp"-->
<!--#include file="myHTMLEncode.asp"-->

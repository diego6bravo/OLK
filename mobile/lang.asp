<%

If myLng = "" Then CheckLanguage

Sub CheckLanguage
	If Request("newLng") <> "" Then
		Session("myLng") = Request("newLng")
		myLng = Request("newLng")
		SetCodePage
	ElseIf Session("myLng") <> "" Then
		myLng = Session("myLng")
	Else
		If Request.cookies("myLng") <> "" Then
			myLng = Request.cookies("myLng")
		Else
			myLng = LCase(Left(Request.ServerVariables("HTTP_ACCEPT_LANGUAGE"),2))
		End If
		If myLng <> "es" and myLng <> "en" and myLng <> "debug" and myLng <> "he" and myLng <> "pt" and myLng <> "fr" and myLng <> "de" and myLng <> "ru" and myLng <> "cs" Then myLng = "en"
		Session("myLng") = myLng
		SetCodePage
	End If
	
	If Request.Cookies("myLng") <> myLng Then
		Response.cookies("myLng").expires = DateAdd("d",60,now())
		Response.cookies("myLng").path = "/"  
		Response.cookies("myLng") = myLng
	End If
End Sub

Sub SetCodePage
	Session.CodePage = 65001
	If myLng = "he" Then Session("rtl") = "rtl/" Else Session("rtl") = ""
	
	Select Case myLng
		Case "en"
			Session("LanID") = 1
		Case "es"
			Session("LanID") = 2
		Case "pt"
			Session("LanID") = 6
		Case "he"
			Session("LanID") = 3
		Case "fr"
			Session("LanID") = 8
		Case "de"
			Session("LanID") = 10
		Case "ru"
			Session("LanID") = 13
		Case "cs"
			Session("LanID") = 14
	End Select
End Sub

Function GetLangErrCode()
	lngRetVal = ""
	Select Case Session("LanID")
		Case 1
			lngRetVal = "EN"
		Case 2
			lngRetVal = "ES_LA"
		Case 3
			lngRetVal = "HE"
		Case 6
			lngRetVal = "PT"
		Case 8
			lngRetVal = "FR"
		Case 10
			lngRetVal = "DE"
		Case 13
			lngRetVal = "RU"
		Case 14
			lngRetVal = "CS"
	End Select
	GetLangErrCode = lngRetVal
End Function
 %><!--#include file="myHTMLEncode.asp"-->
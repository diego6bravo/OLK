<%
If Request("newLng") <> "" Then
	Session("myLng") = Request("newLng")
	myLng = Request("newLng")
ElseIf Session("myLng") <> "" Then
	myLng = Session("myLng")
Else
	If Request.Cookies("myLng").value <> "" Then
		myLng = Request.Cookies("myLng").value
	Else
		myLng = LCase(Left(Request.ServerVariables("HTTP_ACCEPT_LANGUAGE"),2))
	End If
	If myLng <> "es" and myLng <> "en" and myLng <> "debug" and myLng <> "he" and myLng <> "pt" and myLng <> "fr" and myLng <> "de" and myLng <> "ru" Then myLng = "en"
	Session("myLng") = myLng
End If

If Request.Cookies("myLng").value <> myLng Then
	Response.cookies("myLng").expires = DateAdd("d",60,now())
	Response.cookies("myLng").path = "/"  
	Response.cookies("myLng").value = myLng
End If
%>
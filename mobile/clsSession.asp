<!--#include file="langIndex.inc"-->
<%
Class clsSession
	Sub EndDBSession
		Session.Contents.RemoveAll 
	End Sub
	
	Sub StartSession
		strLng = Request.Cookies("myLng")
		If strLng = "" Then strLng = ValidateLang(LCase(Left(Request.ServerVariables("HTTP_ACCEPT_LANGUAGE"),2)))
		Response.cookies("myLng").expires = DateAdd("d",60,now())
		Response.cookies("myLng").path = "/"  
		Response.cookies("myLng") = strLng
		SetCodePage
		
		strAgent = Request.ServerVariables("HTTP_USER_AGENT")

		If InStr(strAgent, "iPad") Then Session("Touch") = True

		Session("Started") = True
	End Sub
	
	Sub CheckSessionStatus
		If Not Session("Started") Then mySession.StartSession
		If Request("newLng") <> "" Then ChangeLanguage(Request("newLng"))
		
		If Session("ID") <> "" Then myApp.CheckLastUpdate
	End Sub
	
	'Language 
	Sub ChangeLanguage(ByVal Language)
		Response.cookies("myLng").expires = DateAdd("d",60,now())
		Response.cookies("myLng").path = "/"  
		Response.cookies("myLng") = ValidateLang(Language)
		SetCodePage
	End Sub
	
	Sub SetCodePage
		Session.CodePage = 65001
		strLng = Request.Cookies("myLng")
		
		For i = 0 to UBound(myLanIndex)
			If myLanIndex(i)(0) = strLng Then
				Session("LanID") = myLanIndex(i)(4)
				Exit For
			End If
		Next
		
		Select Case strLng
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
	
	Function ValidateLang(ByVal Language)
		strLng = Language
		blFound = False
		If strLng = "debug" Then 
			blFound = True
		Else
			For i = 0 to UBound(myLanIndex)
				If myLanIndex(i)(0) = strLng Then
					blFound = True
					Exit For
				End If
			Next
			If Not blFound Then strLng = "en"
		End If
		ValidateLang = strLng
	End Function
	
	Function GetObsLngErrCode
		strLng = ""
		
		For i = 0 to UBound(myLanIndex)
			If myLanIndex(i)(4) = Session("LanID") Then
				strLng = myLanIndex(i)(6)
				Exit For
			End If
		Next
		
		GetObsLngErrCode = strLng
	End Function
	
	'Login
	
	Sub Login(ByVal SessionType)
		Session("Type") = SessionType
	End Sub
	
	Sub LoginDB(ByVal SessionType, ByVal DatabaseID)
		Login(SessionType)
		Session("ID") = DatabaseID
	End Sub
	
	Sub LoginAgent()
		Session("Type") = "AG"
		If Session("UserAccess") = "U" Then
			cmd.CommandText = "DBOLKGetAgentLoginData" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@SlpCode") = Session("vendid")
			set rSlpData = Server.CreateObject("ADODB.RecordSet")
			set rSlpData = cmd.execute()
			Session("MaxLineDiscount") = rSlpData("MaxLineDiscount")
			Session("MaxDocDiscount") = rSlpData("MaxDocDiscount")
		End If
			
	End Sub
	
	Function IsDatabaseLoaded
		IsDatabaseLoaded = Session("ID") <> ""
	End Function
	
	'Session Functions
	
	Function GetObjectContent(ByVal ObjID)
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKGetObjectContent" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@ObjType") = "S"
		cmd("@ObjID") = ObjID
		cmd("@UserType") = userType
		set rObjCont = Server.CreateObject("ADODB.RecordSet")
		rObjCont.open cmd, , 3, 1
		GetObjectContent = CStr(rObjCont("ObjContent"))
	End Function
	
	Dim myCompanyName
	Function GetCompanyName
		If myCompanyName = "" Then
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKGetCmpName" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@LanID") = Session("LanID")
			set rCmpName = Server.CreateObject("ADODB.RecordSet")
			rCmpName.open cmd, , 3, 1
			myCompanyName = CStr(rCmpName(0))
		End If
		GetCompanyName = myCompanyName
	End Function
	
	Dim myAgentName
	Function GetAgentName
		If myAgentName = "" Then
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKGetAgentName" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@LanID") = Session("LanID")
			cmd("@SlpCode") = Session("vendid")
			set rAgentName = Server.CreateObject("ADODB.RecordSet")
			rAgentName.open cmd, , 3, 1
			myAgentName = CStr(rAgentName(0))
		End If
		GetAgentName = myAgentName
	End Function
	
	
	'Session Properties
	
	Function MaxLineDiscount
		MaxLineDiscount = CDbl(Session("MaxLineDiscount"))
	End Function
	
	Function MaxDocDiscount
		MaxDocDiscount = CDbl(Session("MaxDocDiscount"))
	End Function

End Class

%>
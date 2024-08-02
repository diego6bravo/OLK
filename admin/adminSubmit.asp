<!--#include file="chkLogin.asp" -->
<!--#include file="lang/adminSubmit.asp" -->
<!--#include file="adminTradSave.asp"-->
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<!--#include file="lcidReturn.inc"-->

<!--#include file="myHTMLEncode.asp"-->
<% 
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

Dim conn
doWinClose = False
WebErr = ""
Select Case Request("submitCmd")
	Case "adminiPO"
		adminiPO()
	Case "AutGrp"
		adminAutGrp()
	Case "adminDecimals"
		adminDecimals()
	Case "adminPrint"
		adminPrint()
	Case "setLic"
		setLic()
	Case "adminObjConfCols"
		adminObjConfCols()
	Case "adminInformer"
		adminInformer()
	Case "adminCustomSearch"
		adminCustomSearch()
	Case "admCatNav"
		admCatNav()
	Case "admDocFlow"
		admDocFlow()
	Case "admPolls"
		admPolls()
	Case "admNews"
		admNews()
	Case "adminSecRS"
		adminSecRS()
	Case "adminPaySis"
		adminPaySis()
	Case "adminShipSis"
		adminShipSis()
	Case "adminLogos"
		adminLogos()
	Case "adminPrintTitle"
		adminPrintTitle()
	Case "adminCartMore"
		adminCartMore()
	Case "adminUsersLic"
		adminUsersLic()
	Case "menuGroups"
		menuGroups() 
	Case "adminTrad"
		adminTrad()
	Case "adminDefinition"
		adminDefinition()
	Case "adminMyData"
		adminMyData()
	Case "adminBN"
		adminBN()
	Case "adminSecIndex"
		adminSecIndex()
	Case "adminObjs"
		adminObjs()
	Case "adminBatchOpt"
		adminBatchOpt()
	Case "adminCart"
		adminCart()
	Case "alterNames"
		alterNames()
	Case "adminMsg"
		adminMsg()
	Case "adminAcctRejReasons"
		adminAcctRejReasons()
	Case "adminSec"
		adminSec()
	Case "addNews"
		addNews()
	Case "updateNews"
		updateNews()
	Case "adminAnonLogin"
		adminAnonLogin()
	Case "adminAlert"
		adminAlert()
	Case "admininvopt"
		adminInvOpt()
	Case "adminCartOpt"
		adminCartOpt()
	Case "admincrdopt"
		adminCrdOpt()
	Case "admininv"
		adminInv()
	Case "adminnew"
		adminNew()
		response.redirect "adminGeneral.asp"
	Case "adminnote"
		adminNote()
	Case "adminportal"
		adminPortal
	Case "adminCatProp"
		adminCatProp()
	Case "adminNavCat"
		adminNavCat()
	Case "changeDb"
		ID = CInt(Request("dbName"))
		set rs = server.createobject("ADODB.Recordset")
		cmd.ActiveConnection = connCommon 
		cmd.CommandText = "OLKChangeDB"
		cmd.Parameters.Refresh()
		cmd("@ID") = ID
		set rs = cmd.execute()
		If Not rs.Eof Then
			If rs("Verfy") = "Y" Then
				myApp.LoadDBConfigData ID
				response.redirect "admin.asp" 
			Else
				response.redirect "updateDb/updateDb.asp?dbID=" & ID & "&dbName=" & rs("dbName")
			End If
		Else
			Response.Redirect "admin.asp"
		End If
	Case "changeDefDb"
		sql = "update olkAdminLogin set DfltDB = (select dbName from OLKDBA where ID = " & Request("dbID") & ")"
		myApp.ConnectCommon
		conn.execute(sql)
		conn.close
		response.redirect "admin.asp"
	Case "adminPriceCod"
		adminPriceCod()
		response.redirect "adminPriceCod.asp?codType=" & Request("codType")
	Case "adminCUFD"
		adminCUFD()
		response.redirect "adminCUFD.asp?TableID=" & Request("TableID") & "&FieldID=" & Request("LoadFieldID")
	Case "adminCUFDGroups"
		adminCUFDGroups()
		response.redirect "adminCUFD.asp?TableID=" & Request("TableID") & "&#tblGroups"
	Case "adminCatOpt"
		adminCatOpt()
	Case "SingleUserAdd"
		set oLic = server.CreateObject("TM.LicenceConnect.LicenceConnection")
		oLic.LicenceServer = licip
		oLic.LicencePort = licport
		If Request("EMail") <> "" Then EMail = "N'" & Replace(Request("EMail"), "'", "''") & "'" Else EMail = "NULL"
		If Request("EMailInbox") <> "" Then EMailInbox = "Y" Else EMailInbox = "N"
		If Request("Admin") <> "" Then Admin = "Y" Else Admin = "N"
		If Request("chkStatus") = "Y" Then Status = "Y" Else Status = "N"
		sql = "insert OLKAgentsAccess(UserName, Password, Status, EMail, EMailInbox, Admin, ChangePwd, LastUpdate) " & _
		"values(N'" & saveHTMLDecode(Request("UserName"), False) & "', N'" & oLic.GetEncPwd(saveHTMLDecode(Request("Password"), False)) & "', '" & Status & "', " & EMail & ", '" & EMailInbox & "', '" & Admin & "', '" & Request("ChangePwd") & "', getdate()) "

		connCommon.execute(sql)
				
		connCommon.close
		If err.number <> 0 then errN = "&err=" & err.number
		response.redirect "adminSingleAccess.asp?1=1" & errN

	Case "SingleUserRem"
		sql = 	"delete olkagentsaccess where UserName = '" & saveHTMLDecode(Request("UName"), False) & "' " & _
				"delete OLKAgentsAccessDB where UserName = '" & saveHTMLDecode(Request("UName"), False) & "' " & _
				"delete OLKAgentsAccessIPS where UserName = '" & saveHTMLDecode(Request("UName"), False) & "' "

		connCommon.execute(sql)
		
		connCommon.close
		response.redirect "adminSingleAccess.asp"

	Case "SingleUserPwd"
		set oLic = server.CreateObject("TM.LicenceConnect.LicenceConnection")
		oLic.LicenceServer = licip
		oLic.LicencePort = licport
		If Request("ChangePwd") = "Y" Then ChangePwd = "Y" Else ChangePwd = "N"
		sql = "update olkagentsaccess set Password = N'" & oLic.GetEncPwd(saveHTMLDecode(Request("pwd1"), False)) & "', ChangePwd = '" & ChangePwd & "' where username = N'" & saveHTMLDecode(Request("UName"), False) & "'"
		connCommon.execute(sql)
		connCommon.close
		PwdChanged = True
		WinClose("")

	Case "SingleUserUpd"

		If Request("UserName") <> "" Then
			UserName = Split(Request("UserName"), ", ")
			sql = ""
			For i = 0 to UBound(UserName)
				id = i + 1
				If Request("EMail" & id) <> "" Then EMail = "N'" & Replace(Request("EMail" & id), "'", "''") & "'" Else EMail = "NULL"
				If Request("chkEMailInbox" & id) = "Y" Then EMailInbox = "Y" Else EMailInbox = "N"
				If Request("Admin" & id) <> "" Then Admin = "Y" Else Admin = "N"
				If Request("chkStatus" & id) = "Y" Then strStatus = "Y" Else strStatus = "N"
				sql = sql & " update olkagentsaccess set [Status] = '" & strStatus & "', " & _
				"EMail = " & EMail & ", EMailInbox = '" & EMailInbox & "', Admin = '" & Admin & "', LastUpdate = getdate() " & _
				"where UserName = N'" & saveHTMLDecode(UserName(i), False) & "'"
			Next
			connCommon.execute(sql)
		End If
		
		connCommon.close
		response.redirect "adminSingleAccess.asp"
	Case "vUserPwd"
		set oLic = server.CreateObject("TM.LicenceConnect.LicenceConnection")
		oLic.LicenceServer = licip
		oLic.LicencePort = licport
		If Request("ChangePwd") = "Y" Then ChangePwd = "Y" Else ChangePwd = "N"
		sql = "update olkagentsaccess set Password = N'" & oLic.GetEncPwd(saveHTMLDecode(Request("pwd1"), False)) & "', ChangePwd = '" & ChangePwd & "' where username = N'" & Request("UName") & "'"
		conn.execute(sql)
		conn.close
		PwdChanged = True
		WinClose("")
	Case "vUserAdd"
		set oLic = server.CreateObject("TM.LicenceConnect.LicenceConnection")
		oLic.LicenceServer = licip
		oLic.LicencePort = licport
		If Request("AsignedSLP") = "Y" Then AsignedSLP = "Y" Else AsignedSLP = "N"
		If Request("EMail") <> "" Then EMail = "N'" & Replace(Request("EMail"), "'", "''") & "'" Else EMail = "NULL"
		If Request("EMailInbox") = "Y" Then EMailInbox = "Y" Else EMailInbox = "N"
		If Request("Admin") <> "" Then Admin = "Y" Else Admin = "N"
		sql = "insert OLKAgentsAccess(UserName, SlpCode, Password, Access, WhsCode, AsignedSLP, EMail, EMailInbox, Admin, ChangePwd) " & _
		"values(N'" & saveHTMLDecode(Request("UserName"), False) & "', " & Request("SlpCode") & ", N'" & oLic.GetEncPwd(saveHTMLDecode(Request("Password"), False)) & "', '" & Request("Access") & "', N'" & Request("WhsCode") & "', '" & AsignedSLP & "', " & EMail & ", '" & EMailInbox & "', '" & Admin & "', '" & Request("ChangePwd") & "') "

		If Request("rgIndex") <> "" Then
			sql = sql & "insert OLKRGAccess(SlpCode, rgIndex) select " & Request("SlpCode") & ", Value from OLKCommon.dbo.OLKSplit('" & Request("rgIndex") & "', ', ')"
		End If


		On Error Resume Next
		conn.execute(sql)
		
		setActiveMail
				
		conn.close
		If err.number <> 0 then errN = "&err=" & err.number
		response.redirect "adminAgentsAccess.asp?1=1" & errN
	Case "vUserRem"
		sql = 	"delete OLKRGAccess where SlpCode = (select SlpCode from OLKAgentsAccess where UserName = '" & Request("UName") & "') " & _
				"delete olkagentsaccess where UserName = '" & Request("UName") & "' "

		conn.execute(sql)
		
		setActiveMail
		
		conn.close
		response.redirect "adminAgentsAccess.asp"
	Case "vUserUpd"

		sql = "select SlpCode from olkagentsaccess"
		set rs = server.createobject("ADODB.RecordSet")
		set rs = conn.execute(sql)
		sql = ""
		do while not rs.eof
			If Request("AsignedSLP" & rs("SlpCode")) = "Y" Then AsignedSLP = "Y" Else AsignedSLP = "N"
			If Request("EMail" & rs("SlpCode")) <> "" Then EMail = "N'" & Replace(Request("EMail" & rs("SlpCode")), "'", "''") & "'" Else EMail = "NULL"
			If Request("chkEMailInbox" & rs("SlpCode")) = "Y" Then EMailInbox = "Y" Else EMailInbox = "N"
			If Request("Admin" & rs("SlpCode")) <> "" Then Admin = "Y" Else Admin = "N"
			sql = sql & " update olkagentsaccess set access = '" & Request("Access" & rs("SlpCode")) & "', " & _
			"WhsCode = N'" & Request("WhsCode" & rs("SlpCode")) & "', EMailInbox = '" & EMailInbox & "', " & _
			"UserName = N'" & saveHTMLDecode(Request("UserName" & rs("SlpCode")), False) & "', AsignedSLP = '" & AsignedSLP & "', EMail = " & EMail & ", Admin = '" & Admin & "', LastUpdate = getdate() " & _
			"where slpcode = " & rs("SlpCode")
		rs.movenext
		loop
		conn.execute(sql)
		
		setActiveMail
		
		conn.close
		response.redirect "adminAgentsAccess.asp"
	Case "delNews"
		sql = "update olkNews set Status = 'D' where newsIndex = " & Request("newsIndex")

		conn.execute(sql)
		conn.close
		response.redirect "adminNews.asp"
	Case "adminDocFlow"
		adminDocFlow()
	Case "delDocFlow"
		set rs = server.createobject("ADODB.RecordSet")
		sql = "update OLKUAF set Active = 'D' where FlowID = " & Request("FlowID")
		conn.execute(sql)
		conn.close
		response.redirect "adminDocFlow.asp"
	Case "adminPwd"
		adminAPwd()
	Case "adminLang"

		sql = "delete OLKDisLng "
		LanID = Split(Request("LanID"), ", ")
		For i = 0 to UBound(LanID)
			If Request("LanID" & LanID(i)) = "" Then
				sql = sql & "insert OLKDisLng(LanID) values(" & LanID(i) & ") "
			End If
		Next
		
		conn.execute(sql)
		setActiveMail

		conn.close
		response.redirect "adminLanguages.asp"
End Select

Sub adminAutGrp
	Select Case Request("cmd")
		Case "uGrp"
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKSaveAutGrpData" & Session("ID")
			cmd.Parameters.Refresh()
			
			arrGrp = Split(Request("GrpID"), ", ")
			For i = 0 to UBound(arrGrp)
				GrpID = arrGrp(i)
				cmd("@GrpID") = GrpID
				cmd("@GroupName") = Request("GrpName" & GrpID)
				If Request("Branch" & GrpID) = "Y" Then cmd("@FilterBranch") = "Y" Else cmd("@FilterBranch") = "N"
				cmd.execute()
			Next
		Case "delGrp"
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKDelAutGrp" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@GrpID") = Request("delIndex")
			cmd.execute()
		Case "AutGrpData"
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKSaveAutGrpData" & Session("ID")
			cmd.Parameters.Refresh()
			If Request("GrpID") <> "" Then cmd("@GrpID") = Request("GrpID")
			cmd("@GroupName") = Request("GroupName")
			If Request("chkBranch") = "Y" Then cmd("@FilterBranch") = "Y" Else cmd("@FilterBranch") = "N"
			cmd.execute()
			
			GrpID = cmd("@GrpID")
			
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKSaveSlpFilterAutGrp" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@GrpID") = GrpID
			
			arrSlp = Split(Request("SlpCode"), ", ")
			For i = 0 to UBound(arrSlp)
				SlpCode = arrSlp(i)
				Op = Request("Op" & SlpCode)
				Ordr = Request("Order" & SlpCode)
				cmd("@SlpCode") = SlpCode
				cmd("@Op") = Op
				cmd("@Ordr") = Ordr
				cmd.execute()
			Next
			
			If Request("delID") <> "" Then
				set cmd = Server.CreateObject("ADODB.Command")
				cmd.ActiveConnection = connCommon
				cmd.CommandType = &H0004
				cmd.CommandText = "DBOLKDelSlpFilterAutGrp" & Session("ID")
				cmd.Parameters.Refresh()
				cmd("@GrpID") = GrpID
				arrDel = Split(Request("delID"), ", ")
				For i = 0 to UBound(arrDel)
					cmd("@SlpCode") = arrDel(i)
					cmd.execute()
				Next
			End If
						
			If Request("btnApply") <> "" Then Response.Redirect "adminAutGrp.asp?GrpID=" & GrpID
	End Select
	Response.Redirect "adminAutGrp.asp"
End Sub

Sub adminDecimals
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKSetDecSettingsData" & Session("ID")
	cmd.Parameters.Refresh()
	If Request("AlterQtyDec") <> "" Then cmd("@AlterQtyDec") = Request("AlterQtyDec")
	If Request("AlterPriceDec") <> "" Then cmd("@AlterPriceDec") = Request("AlterPriceDec")
	If Request("AlterPercentDec") <> "" Then cmd("@AlterPercentDec") = Request("AlterPercentDec")
	If Request("AlterMeasureDec") <> "" Then cmd("@AlterMeasureDec") = Request("AlterMeasureDec")
	If Request("AlterSumDec") <> "" Then cmd("@AlterSumDec") = Request("AlterSumDec")
	If Request("AlterRateDec") <> "" Then cmd("@AlterRateDec") = Request("AlterRateDec")
	cmd.execute()	
	
	myApp.LoadDecSettings
	myApp.ResetLastUpdate

	
	Response.Redirect "adminCustDec.asp"
End Sub

Private Sub setLic

	set oLic = server.CreateObject("TM.LicenceConnect.LicenceConnection")
	
		with olic
	        .LicenceServer = licip
	        .LicencePort = licport
	        
	        isAlive = oLic.IsAlive
	        
			if isAlive then 
				arrUsers = Split(Request("ID"), ", ")
				For i = 0 to UBound(arrUsers)
					SlpCode = arrUsers(i)
					userName = Request("UserName" & SlpCode)
					agentLic = Request("ChkAgent" & SlpCode) = "Y"
					mobileLic = Request("ChkMobile" & SlpCode) = "Y"
					amLic = Request("ChkAM" & SlpCode) = "Y"
					oLic.SetLicUsers 51, 0, userName, agentLic
					oLic.SetLicUsers 52, 0, userName, mobileLic
					oLic.SetLicUsers 53, 0, userName, amLic
				Next
			end if 
		end with 
	
	Response.Redirect "adminLicInf.asp"

End Sub

Sub adminiPO
	If Request("ListNumID") <> "" Then
			
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKSetiPOPriceList" & Session("ID")
		cmd.Parameters.Refresh()
		
		arr = Split(Request("ListNumID"), ", ")
		For i = 0 to UBound(arr)
			ListNum = arr(i)
			cmd("@ListNum") = ListNum
			If Request("chkListActive" & ListNum) = "Y" Then cmd("@Active") = "Y" Else cmd("@Active") = "N"
			cmd("@Ordr") = Request("Order" & ListNum)
			cmd.execute()
		Next
		
	End If
	
	If Request("rowID") <> "" Then
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKSetiPOField" & Session("ID")
		cmd.Parameters.Refresh()
		
		arr = Split(Request("rowID"), ", ")
		For i = 0 to UBound(arr)
			rowID = arr(i)
			cmd("@RowIndex") = rowID
			rowID = Replace(rowID, "-", "_")
			If Request("chkRowActive" & rowID) = "Y" Then cmd("@Active") = "Y" Else cmd("@Active") = "N"
			cmd.execute()
		Next
	End If

	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKSetiPOQuery" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@Sizes") = Request("qrySizes")
	cmd("@SubCategory") = Request("qrySubCategory")
	cmd("@SubSubCategory") = Request("qrySubSubCategory")
	cmd("@Origin") = Request("qryOrigin")
	cmd("@Brand") = Request("qryBrand")
	cmd("@Packing") = Request("qryPacking")
	cmd("@Composition") = Request("qryComposition")
	cmd("@Season") = Request("qrySeason")
	cmd.execute()
	
	GenMyQuery "ItemRepAll"
	
	Response.Redirect "adminiPO.asp"	
End Sub

Sub adminPrint
	object = CInt(Request("ObjectCode"))
	
	Select Case Request("cmd")
		Case "U"
			If Request("SecID") <> "" Then
				set cmd = Server.CreateObject("ADODB.Command")
				cmd.ActiveConnection = connCommon
				cmd.CommandType = &H0004
				cmd.CommandText = "DBOLKSetObjectPrintData" & Session("ID")
				cmd.Parameters.Refresh()
				cmd("@ObjectCode") = object
				
				arrSecID = Split(Request("SecID"), ", ")
				
				For i = 0 to UBound(arrSecID)
					secID = arrSecID(i)
					
					LinkData = Request("LinkData" & secID)
					Order = CInt(Request("Order" & secID))
					If Request("Active" & secID) = "Y" Then Active = "Y" Else Active = "N"
					
					cmd("@SecID") = CInt(Replace(secID, "_", "-"))
					cmd("@LinkData") = LinkData
					cmd("@LinkName") = Request("LinkName" & secID)
					cmd("@Order") = Order
					cmd("@Active") = Active
					
					cmd.execute()
				Next
			End If
		Case "D"
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKDelObjectPrint" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@ObjectCode") = object
			cmd("@SecID") = Request("delID")
			cmd.execute()
		Case "A"
			SecID = CInt(Request("SecID"))
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKAddObjectPrintData" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@ObjectCode") = object
			cmd("@SecID") = SecID
			If SecID < 0 Then cmd("@LinkType") = Request("userType")
			cmd.execute()
	End Select
	
	Response.Redirect "adminObjPrint.asp?object=" & object
End Sub

Private Sub adminInformer
	Select Case Request("cmd")
		Case "update"
			arrType = Split(Request("Type"), ", ")
			arrID = Split(Request("ID"), ", ")
			
			For i = 0 to UBound(arrType)
				strType = arrType(i)
				strID = arrID(i)
				
				myID = strType & strID
				
				If Request("RowActive" & myID) = "Y" Then Active = "Y" Else Active = "N"
				If Request("RowHideNull" & myID) = "Y" Then HideNull = "Y" Else HideNull = "N"
				If Request("RowAlign" & myID) <> "" Then Align = "'" & Request("RowAlign" & myID) & "'" Else Align = "NULL"
				
				sql = "update OLKInformer set Ordr = " & Request("RowOrder" & myID) & ", Active = '" & Active & "'"
				
				If strType = "U" Then
					sql = sql & ", [Name] = N'" & saveHTMLDecode(Request("RowName" & myID), False) & "', HideNull = '" & HideNull & "', Align = " & Align & " "
				End If
				
				sql = sql & " where [Type] = '" & strType & "' and [ID] = " & strID
				
				conn.execute(sql)
				
			Next
		Case "del"
			sql = "delete OLKInformer where Type = 'U' and ID = " & Request("ID")
			conn.execute(sql)
		Case "save"
			ID = Request("ID")
			
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKSaveInformerData" & Session("ID")
			cmd.Parameters.Refresh()
			If ID <> "" Then cmd("@ID") = ID
			cmd("@Name") = Request("RowName")
			If Request("RowAlign") <> "" Then cmd("@Align") = Request("RowAlign")
			cmd("@Query") = Request("RowQuery")
			If Request("RowHideNull") = "Y" Then cmd("@HideNull") = "Y" Else cmd("@HideNull") = "N"
			If Request("RowActive") = "Y" Then cmd("@Active") = "Y" Else cmd("@Active") = "N"
			If Request("RowReport") <> "" Then cmd("@rsIndex") = Request("RowReport")
			cmd("@Ordr") = Request("RowOrder")
			cmd.execute()
			
			If ID = "" Then
				ID = cmd.Parameters(0).value
				
				If Request("rowNameTrad") <> "" Then
					SaveNewTrad Request("rowNameTrad"), "Informer", "ID", "AlterName", ID 
				End If
				
				If Request("varQueryDef") <> "" Then
					SaveNewDef Request("RowQueryDef"), ID
				End If

			End If
			
			If Request("btnApply") <> "" Then Response.Redirect "adminInformerEdit.asp?ID=" & ID
			
	End Select
	
	Response.Redirect "adminInformer.asp"
End Sub

Private Sub admCatNav()
	If Request("NavIndex") <> "" Then

		
		arrNav = Split(Request("NavIndex"), ", ")
		For i = 0 to UBound(arrNav)
			NavIndex = arrNav(i)
			If Request("Active" &  NavIndex) = "Y" Then Active = "Y" Else Active = "N"
			sql = "update OLKCatNav set Active = '" & Active & "' where NavIndex = " & NavIndex
			conn.execute(sql)
		Next
	End If

	Response.Redirect "adminCatNav.asp"
End Sub

Private Sub admDocFlow()
	If Request("FlowID") <> "" Then


		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKGenQry" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@Type") = "DF"

		arrFlow = Split(Request("FlowID"), ", ")
		For i = 0 to UBound(arrFlow)
			FlowID = arrFlow(i)
			If Request("Active" &  FlowID) = "Y" Then Active = "Y" Else Active = "N"
			sql = "update OLKUAF set Active = '" & Active & "', [Order] = " & Request("Order" & FlowID) & " where FlowID = " & FlowID
			conn.execute(sql)
			
			If Request("Active" & FlowID) = "Y" Then
				cmd("@ID") = FlowID
				cmd.execute()
			End If
		Next
	End If
	
	Response.Redirect "adminDocFlow.asp"
End Sub

Private Sub admPolls()
	If Request("pollIndex") <> "" Then

		
		arrPoll = Split(Request("pollIndex"), ", ")
		For i = 0 to UBound(arrPoll)
			pollIndex = arrPoll(i)
			If Request("Status" &  pollIndex) = "O" Then Status = "O" Else Status = "C"
			sql = "update OLKPoll set pollStatus = '" & Status & "' where pollIndex = " & pollIndex
			conn.execute(sql)
		Next
	End If
	
	Response.Redirect "adminPolls.asp"
End Sub

Private Sub admNews()
	If Request("newsIndex") <> "" Then

		
		arrNews = Split(Request("newsIndex"), ", ")
		For i = 0 to UBound(arrNews)
			newsIndex = arrNews(i)
			If Request("Status" &  newsIndex) = "A" Then Status = "A" Else Status = "N"
			sql = "update OLKNews set Status = '" & Status & "' where newsIndex = " & newsIndex
			conn.execute(sql)
		Next
	End If
	
	Response.Redirect "adminNews.asp"
End Sub

Private Sub adminSecRS()
	set rs = Server.CreateObject("ADODB.RecordSet")
	
	Select Case Request("cmd")
		Case "edit"
			If Request("LineID") = "" Then
				sql = "declare @LineID int set @LineID = IsNull((select Max(LineID)+1 from OLKSectionsRS where SecType = 'U' and SecID = " & Request("SecID") & "), 0) " & _
					"select @LineID LineID " & _
					"insert OLKSectionsRS(SecType, SecID, LineID, Name, Query) " & _
					"values('U', " & Request("SecID") & ", @LineID, N'" & saveHTMLDecode(Request("Name"), False) & "', N'" & saveHTMLDecode(Request("Query"), False) & "') "
				set rs = conn.execute(sql)
				LineID = rs("LineID")

				If Request("queryDef") <> "" Then
					SaveNewDef Request("queryDef"), Request("SecID") & LineID
				End If
			Else
				sql = "update OLKSectionsRS set Name = N'" & saveHTMLDecode(Request("Name"), False) & "', Query = N'" & saveHTMLDecode(Request("Query"), False) & "' " & _
					"where SecType = 'U' and SecID = " & Request("SecID") & " and LineID = " & Request("LineID")
				conn.execute(sql)
				LineID = Request("LineID")
			End If
			
			If Request("btnApply") <> "" Then Response.Redirect "adminSecRS.asp?SecID=" & Request("SecID") & "&LineID=" & LineID
		Case "del"
			sql = "delete OLKSectionsRS where SecType = 'U' and SecID = " & Request("SecID") & " and LineID = " & Request("LineID")
			conn.execute(sql)
	End Select
	
	Response.Redirect "adminSecRS.asp?SecID=" & Request("SecID")
End Sub

Private Sub adminCustomSearch()
	ObjID = Request("ObjID")
	Select Case Request("cmd")
		Case "Prop"
			sql = ""
			ID = Request("ID")
			
			For i = 1 to 64
				If Request("qryGroup" & i) = "Y" Then Active = "Y" Else Active = "N"
				Ordr = Request("Ordr" & i)
				sql = sql & " update OLKCustomSearchProp set Active = '" & Active & "', Ordr = " & Ordr & " where ObjectCode = " & ObjID & " and ID = " & ID & " and PropID = " & i & " "
			Next
			conn.execute(sql)
			
			WinClose("")
		Case "u"
			arrID = Split(Request("ID"), ", ")
			For i = 0 to UBound(arrID)
				ID = Replace(arrID(i), "-", "_")
				Name = saveHTMLDecode(Request("rowName" & ID), False)
				Ordr = Request("RowOrder" & ID)
				If Request("rowActive" & ID) = "Y" Then Status = "Y" Else Status = "N"
				sql = "update OLKCustomSearch set Name = N'" & Name & "', Ordr = " & Ordr & ", Status = '" & Status & "' where ObjectCode = " & ObjID & " and ID = " & arrID(i)
				conn.execute(sql)
			Next
		Case "del"
			sql = "update OLKCustomSearch set Status = 'D' where ObjectCode = " & ObjID & " and ID = " & Request("ID")
			conn.execute(sql)
		Case "save"
			ID = Request("ID")
			Name = saveHTMLDecode(Request("searchName"), False)
			If Request("chkIgnoreGeneralFilter") = "Y" Then IgnoreGeneralFilter = "Y" Else IgnoreGeneralFilter = "N"
			Query = saveHTMLDecode(Request("txtQry"), False)
			CatType = Request("CatType")
			If Request("Order1") <> "" Then Order1 = "N'" & Request("Order1") & "'" Else Order1 = "NULL"
			If Request("Order2") <> "" Then Order2 = "'" & Request("Order2") & "'" Else Order2 = "NULL"
			Ordr = Request("RowOrder")
			If Request("chkActive") = "Y" Then Status = "Y" Else Status = "N"
			
			If ID <> "" Then
				sql = "update OLKCustomSearch set Name = N'" & Name & "', IgnoreGeneralFilter = '" & IgnoreGeneralFilter & "', Query = N'" & Query & "', CatType = '" & CatType & "', Order1 = " & Order1 & ", Order2 = " & Order2 & ", " & _
						"Ordr = " & Ordr & ", Status = '" & Status & "' where ID = " & ID & " and ObjectCode = " & ObjID
				conn.execute(sql)
			Else
				sql = "declare @ID int set @ID = IsNull((select Max(ID)+1 from OLKCustomSearch where ObjectCode = " & Request("ObjID") & "), 0) select @ID ID " & _
						"insert OLKCustomSearch(ObjectCode, ID, Name, IgnoreGeneralFilter, Query, CatType, Order1, Order2, Status, Ordr) " & _
						"values(" & ObjID & ", @ID, N'" & Name & "', '" & IgnoreGeneralFilter & "', N'" & Query & "', '" & CatType & "', " & Order1 & ", " & Order2 & ", '" & Status & "', " & Ordr & ")"
				set rs = Server.CreateObject("ADODB.RecordSet")
				set rs = conn.execute(sql)
				ID = rs(0)
				rs.close
			End If
			
				
			If Request("varNameTrad") <> "" Then
				SaveNewTrad Request("varNameTrad"), "CustomSearch", "ObjectCode,ID", "AlterName", ObjID & "," & ID 
			End If
			
			If Request("varQueryDef") <> "" Then
				SaveNewDef Request("varQueryDef"), ObjID & ID
			End If
			
			sql = "declare @ObjectCode int set @ObjectCode = " & ObjID & " declare @ID int set @ID = " & ID & " " & _
				"declare @Session nvarchar(20) set @Session = '" & Request("Session") & "' " & _
				"delete OLKCustomSearchSession where ObjectCode = @ObjectCode and ID = @ID " & _
				"If @Session <> '' Begin " & _
				"	insert OLKCustomSearchSession(ObjectCode, ID, SessionID) " & _
				"	select @ObjectCode , @ID, Value from OLKCommon.dbo.OLKSplit(@Session, ', ') " & _
				"End "
			conn.execute(sql)
			
			If Request("btnApply") <> "" Then Response.Redirect "adminCustomSearchEdit.asp?ID=" & ID & "&ObjID=" & ObjID
		Case "uVars"
			ID = Request("ID")
			
			If Request("VarID") <> "" Then
				arrID = Split(Request("VarID"), ", ")
				For i = 0 to UBound(arrID)
					VarID = arrID(i)
					Name = saveHTMLDecode(Request("varName" & VarID), False)
					Ordr = Request("Ordr" & VarID)
					sql = "update OLKCustomSearchVars set Name = N'" & Name & "', Ordr = " & Ordr & " where ObjectCode = " & ObjID & " and ID = " & ID & " and VarID = " & VarID
					conn.execute(sql)
				Next
			End If
			
			Response.Redirect "adminCustomSearchEdit.asp?ID=" & ID & "&ObjID=" & ObjID
		Case "delVar"
			ID = Request("ID")
			
			sql = "declare @ObjectCode int set @ObjectCode = " & ObjID & " " & _
					"declare @ID int set @ID = " & ID & " declare @VarID int set @VarID = " & Request("VarID") & " " & _
					"declare @delProp char(1) set @delProp = Case When (select Variable from OLKCustomSearchVars where ObjectCode = @ObjectCode and ID = @ID and VarID = @VarID) in ('ItmProp', 'BPProp') Then 'Y' Else 'N' End " & _
					"delete OLKCustomSearchVars where ObjectCode = @ObjectCode and ID = @ID and VarID = @VarID " & _
					"delete OLKCustomSearchVarsAlterNames where ObjectCode = @ObjectCode and ID = @ID and VarID = @VarID " & _
					"delete OLKCustomSearchVarsVals where ObjectCode = @ObjectCode and ID = @ID and VarID = @VarID " & _
					"delete OLKCustomSearchVarsValsAlterNames where ObjectCode = @ObjectCode and ID = @ID and VarID = @VarID " & _
					"delete OLKCustomSearchVarsBase where ObjectCode = @ObjectCode and ID = @ID and VarID = @VarID " & _
					"If @delProp = 'Y' begin delete OLKCustomSearchPropOrder where ObjectCode = @ObjectCode and ID = @ID End " 
			conn.execute(sql)
			
			Response.Redirect "adminCustomSearchEdit.asp?ObjID=" & ObjID & "&ID=" & ID
		Case "editVar"
			ID = Request("ID")
			VarID = Request("VarID")
			Name = Request("varName")
			Variable = Request("varVar")
			myType = Request("varType")
			DataType = Request("varDataType")
			If Request("varQuery") <> "" Then Query = "N'" & saveHTMLDecode(Request("varQuery"), False) & "'" Else Query = "NULL"
			If Request("varQueryField") <> "" Then QueryField = "N'" & saveHTMLDecode(Request("varQueryField"), False) & "'" Else QueryField = "NULL"
			If Request("varMaxChar") <> "" Then MaxChar = Request("varMaxChar") Else MaxChar= "NULL"
			If Request("varNotNull") = "Y" Then NotNull = "Y" Else NotNull = "N"
			DefVars = Request("varQueryBy")
			DefValBy = Request("varDefBy")
			
			If DefValBy = "V" and DataType <> "datetime" Then
				DefValValue = "N'" & saveHTMLDecode(Request("varDefValValue"), False) & "'"
				DefValDate = "NULL"
			ElseIf DefValBy = "V" and DataType = "datetime" Then
				DefValValue = "NULL"
				DefValDate = "Convert(datetime,'" & SaveSqlDate(Request("varDefValDate")) & "',120)"
			ElseIf DefValBy = "Q" Then
				DefValValue = "N'" & saveHTMLDecode(Request("varDefValQuery"), False) & "'"
				DefValDate = "NULL"
			Else
				DefValValue = "NULL"
				DefValDate = "NULL"
			End If
			 
			Ordr = Request("Ordr")
			
			If VarID <> "" Then
				sql = 	"update OLKCustomSearchVars set Name = N'" & Name & "', Variable = N'" & Variable & "', [Type] = '" & myType & "', DataType = '" & DataType & "', " & _
						"Query = " & Query & ", QueryField = " & QueryField & ", MaxChar = " & MaxChar & ", NotNull = '" & NotNull & "', DefVars = '" & DefVars & "', " & _
						"DefValBy = '" & DefValBy & "', DefValValue = " & DefValValue & ", DefValDate = " & DefValDate & ", Ordr = " & Ordr & " " & _
						"where ObjectCode = " & ObjID & " and ID = " & ID & " and VarID = " & VarID
				conn.execute(sql)
			Else
				sql = "declare @VarID int set @VarID = IsNull((select Max(VarID)+1 from OLKCustomSearchVars where ObjectCode = " & ObjID & " and ID = " & ID & "), 0) select @VarID VarID " & _
						"insert OLKCustomSearchVars(ObjectCode, ID, VarID,[Name], Variable, [Type], DataType, Query, QueryField, MaxChar, NotNull, DefVars, DefValBy, DefValValue, DefValDate, Ordr) " & _
						"values(" & ObjID & ", " & ID & ", @VarID, N'" & Name & "', N'" & Variable & "', '" & myType & "', '" & DataType & "', " & Query & ", " & QueryField & ", " & MaxChar & ", " & _
						"'" & NotNull & "', '" & DefVars & "', '" & DefValBy & "', " & DefValValue & ", " & DefValDate & ", " & Ordr & ") "
						
				set rs = Server.CreateObject("ADODB.RecordSet")
				set rs = conn.execute(sql)
				
				VarID = rs(0)
				
				set rs = nothing
				
				If Request("varNameTrad") <> "" Then
					SaveNewTrad Request("varNameTrad"), "CustomSearchVars", "ObjectCode,ID,VarID", "AlterName", ObjID & "," & ID & "," & VarID
				End If
				
				If Request("varQueryDef") <> "" Then
					SaveNewDef Request("varQueryDef"), ObjID & ID & VarID
				End If
				
				If Request("varDefValueDef") <> "" Then
					SaveNewDef Request("varDefValueDef"), ObjID & ID & VarID
				End If
			End If
			
			If Request("varQueryBy") = "F" Then
				sql = "declare @ObjectCode int set @ObjectCode = " & ObjID & " declare @ID int set @ID = " & ID & " declare @VarID int set @VarID = " & VarID & " " & _
						"delete OLKCustomSearchVarsVals where ObjectCode = @ObjectCode and ID = @ID and VarID = @VarID declare @ValID int "
	
				ArrVal = Split(Request("varQuery"),VbCrLf)
				for i = 0 to UBound(ArrVal)
					ArrVal2 = Split(ArrVal(i),",")
					sql = sql & "set @ValID = IsNull((select Max(ValID) + 1 from OLKCustomSearchVarsVals where ObjectCode = @ObjectCode and ID = @ID and VarID = @VarID), 0) " & _
								"insert OLKCustomSearchVarsVals(ObjectCode, ID, VarID, ValID, Value, Description) " & _
								"values(@ObjectCode, @ID, @VarID, @ValID, N'" & saveHTMLDecode(ArrVal2(0), False) & "', N'" & saveHTMLDecode(ArrVal2(1), False) & "') "
				next
				conn.execute(sql)
			End If
			
			sql = "delete OLKCustomSearchVarsBase where ObjectCode = " & ObjID & " and ID = " & ID & " and VarID = " & VarID & " "
			
			If Request("baseVar") <> "" Then 
				sql = sql & "insert OLKCustomSearchVarsBase(ObjectCode, ID, VarID, BaseID) " & _
							"select " & ObjID & ", " & ID & ", " & VarID & ", [Value] " & _
							"from OLKCommon.dbo.OLKSplit('" & Request("baseVar") & "', ', ') "
			End If
			conn.execute(sql)
			
			If Request("btnApply") <> "" Then
				Response.Redirect "adminCustomSearchEdit.asp?ObjID=" & ObjID & "&ID=" & ID & "&VarID=" & VarID
			Else
				Response.Redirect "adminCustomSearchEdit.asp?ObjID=" & ObjID & "&ID=" & ID
			End If
		Case "addSysVar"
			ID = Request("ID")
			myType = Request("varType")
			
			sql = "declare @ObjectCode int set @ObjectCode = " & ObjID & " declare @ID int set @ID = " & ID  & " " & _
					"declare @VarID int set @VarID = IsNull((select Max(VarID)+1 from OLKCustomSearchVars where ObjectCode = @ObjectCode and ID = @ID), 0) " & _
					"declare @Ordr int set @Ordr = IsNull((select Max(Ordr)+1 from OLKCustomSearchVars where ObjectCode = @ObjectCode and ID = @ID), 0) " & _
					"insert OLKCustomSearchVars(ObjectCode, ID, VarID, [Name], Variable, [Type], DataType, NotNull, DefVars, DefValBy, Ordr) " & _
					"select @ObjectCode, @ID, @VarID, N'" & myType & "', N'" & myType & "', 'S', 'S', 'N', 'Q', 'N', @Ordr " & _
					"where not exists(select '' from OLKCustomSearchVars where ObjectCode = @ObjectCode and ID = @ID and [Type] = 'S' and Variable = N'" & myType & "') "
			
			If myType = "ItmProp" or myType = "BPProp" Then
				sql = sql & 	"declare @i int set @i = 1 " & _  
								"while @i <= 64 begin " & _  
								"	insert OLKCustomSearchProp(ObjectCode, ID, PropID, Active, Ordr) " & _  
								"	values(@ObjectCode, @ID, @i, 'Y', @i) " & _  
								"	set @i = @i + 1 " & _  
								"End " 
				
			End If
			
			conn.execute(sql)
			
			Response.Redirect "adminCustomSearchEdit.asp?ObjID=" & ObjID & "&ID=" & ID
		Case "restore"
			ID = Request("ID")
			
			sql = "declare @ObjectCode int set @ObjectCode = " & ObjID & " declare @ID int set @ID = " & ID & " " & _
					"declare @LanID int set @LanID = " & Session("LanID") & " " & _
					"delete OLKCustomSearch where ObjectCode = @ObjectCode and ID = @ID " & _
					"delete OLKCustomSearchAlterNames where ObjectCode = @ObjectCode and ID = @ID " & _
					"delete OLKCustomSearchVars where ObjectCode = @ObjectCode and ID = @ID " & _
					"insert OLKCustomSearch(ObjectCode, ID, Name, Query, CatType, Order1, Order2, Status, Ordr) " & _
					"select T0.ObjectCode, T0.ID, T1.AlterName, T0.Query, T0.CatType, T0.Order1, T0.Order2, T0.Status, T0.Ordr " & _
					"from OLKCommon..OLKCustomSearch T0 " & _
					"inner join OLKCommon..OLKCustomSearchAlterNames T1 on T1.ObjectCode = T0.ObjectCode and T1.ID = T0.ID " & _
					"where T1.LanID = @LanID and T0.ObjectCode = @ObjectCode and T0.ID = @ID " & _
					"insert OLKCustomSearchAlterNames(LanID, ObjectCode, ID, AlterName) " & _
					"select T0.LanID, T0.ObjectCode, T0.ID, T0.AlterName " & _
					"from OLKCommon..OLKCustomSearchAlterNames T0 " & _
					"where LanID <> @LanID and T0.ObjectCode = @ObjectCode and T0.ID = @ID " & _
					"insert OLKCustomSearchVars(ObjectCode, ID, VarID, Name, Variable, Type, DataType, Query, QueryField, MaxChar, NotNull, DefVars, DefValBy, DefValValue, DefValDate, Ordr) " & _
					"select ObjectCode, ID, VarID, Name, Variable, Type, DataType, Query, QueryField, MaxChar, NotNull, DefVars, DefValBy, DefValValue, DefValDate, Ordr " & _
					"from OLKCommon..OLKCustomSearchVars T0 " & _
					"where T0.ObjectCode = @ObjectCode and T0.ID = @ID "
					
			conn.execute(sql)
			
			Response.Redirect "adminCustomSearchEdit.asp?ObjID=" & ObjID & "&ID=" & ID

	End Select
	
	conn.close
	If Not doWinClose Then Response.Redirect "adminCustomSearch.asp?ObjID=" & ObjID
End Sub

Private Sub adminPaySis()
	
	If Request("chkActive") = "Y" Then Active = "Y" Else Active = "N"
	sql = "declare @PayTypeID int set @PayTypeID = " & Request("PayTypeID") & " " & _
			"if not exists(select '' from OLKPayment where PayTypeID = @PayTypeID) begin " & _
			"	insert OLKPayment(PayTypeID, Active) values(@PayTypeID, '" & Active & "') " & _
			"end else begin " & _
			"	update OLKPayment set Active = '" & Active & "' where PayTypeID = @PayTypeID " & _
			"end "
	
	If Request("FieldID") <> "" Then
		FieldID = Split(Request("FieldID"), ", ")
		sql = sql & "declare @FieldID int "
		For i = 0 to UBound(FieldID)
			If Request("Value" & i) <> "" Then Value = "N'" & Request("Value" & i) & "'" Else Value = "NULL"
			sql = sql & "set @FieldID = " & FieldID(i) & " " & _
			"if not exists(select '' from OLKPaymentSettings where PayTypeID = @PayTypeID and FieldID = @FieldID) begin " & _
			"	insert OLKPaymentSettings(PayTypeID, FieldID, Value) values(@PayTypeID, @FieldID, " & Value & ") " & _
			"end else begin " & _
			"	update OLKPaymentSettings set Value = " & Value & " where PayTypeID = @PayTypeID and FieldID = @FieldID " & _
			"end "
		Next
	End If
	
	If Request("CurrID") <> "" Then
		CurrID = Split(Request("CurrID"), ", ")
		sql = sql & "declare @CurrID int "
		For i = 0 to UBound(CurrID)
			If Request("CurMatch" & CurrID(i)) <> "" Then Match = "N'" & Request("CurMatch" & CurrID(i)) & "'" Else Match = "NULL"
			sql = sql & "set @CurrID = " & CurrID(i) & " " & _
				"if not exists(select '' from OLKPaymentCurMatch where PayTypeID = @PayTypeID and CurrID = @CurrID) begin " & _
				"	insert OLKPaymentCurMatch(PayTypeID, CurrID, Match) values(@PayTypeID, @CurrID, " & Match & ") " & _
				"end else begin " & _
				"	update OLKPaymentCurMatch set Match = " & Match & " where PayTypeID = @PayTypeID and CurrID = @CurrID " & _
				"end "
		Next
	End If
	
	If Request("CardID") <> "" Then
		CardID = Split(Request("CardID"), ", ")
		sql = sql & "declare @CardID int "
		For i = 0 to UBound(CardID)
			If Request("CardMatch" & CardID(i)) <> "" Then Match = Request("CardMatch" & CardID(i)) Else Match = "NULL"
			sql = sql & "set @CardID = " & CardID(i) & " " & _
				"if not exists(select '' from OLKPaymentCardMatch where PayTypeID = @PayTypeID and CardID = @CardID) begin " & _
				"	insert OLKPaymentCardMatch(PayTypeID, CardID, Match) values(@PayTypeID, @CardID, " & Match & ") " & _
				"end else begin " & _
				"	update OLKPaymentCardMatch set Match = " & Match & " where PayTypeID = @PayTypeID and CardID = @CardID " & _
				"end "
		Next
	End If
	
	conn.execute(sql)
	
	conn.close
	
	If Request("btnApply") <> "" Then
		Response.Redirect "adminPaySis.asp?PayTypeID=" & Request("PayTypeID")
	Else
		Response.Redirect "adminPaySis.asp"
	End If
End Sub

Private Sub adminShipSis()

	
	If Request("chkActive") = "Y" Then Active = "Y" Else Active = "N"
	sql = "declare @ShipTypeID int set @ShipTypeID = " & Request("ShipTypeID") & " " & _
			"if not exists(select '' from OLKShipment where ShipTypeID = @ShipTypeID) begin " & _
			"	insert OLKShipment(ShipTypeID, Active) values(@ShipTypeID, '" & Active & "') " & _
			"end else begin " & _
			"	update OLKShipment set Active = '" & Active & "' where ShipTypeID = @ShipTypeID " & _
			"end "
	
	If Request("FieldID") <> "" Then
		FieldID = Split(Request("FieldID"), ", ")
		sql = sql & "declare @FieldID int "
		For i = 0 to UBound(FieldID)
			If Request("Value" & i) <> "" Then Value = "N'" & Request("Value" & i) & "'" Else Value = "NULL"
			sql = sql & "set @FieldID = " & FieldID(i) & " " & _
			"if not exists(select '' from OLKShipmentSettings where ShipTypeID = @ShipTypeID and FieldID = @FieldID) begin " & _
			"	insert OLKShipmentSettings(ShipTypeID, FieldID, Value) values(@ShipTypeID, @FieldID, " & Value & ") " & _
			"end else begin " & _
			"	update OLKShipmentSettings set Value = " & Value & " where ShipTypeID = @ShipTypeID and FieldID = @FieldID " & _
			"end "
		Next
	End If
	
	If Request("LenID") <> "" Then
		LenID = Split(Request("LenID"), ", ")
		sql = sql & "declare @LenID int "
		For i = 0 to UBound(LenID)
			If Request("LenMatch" & LenID(i)) <> "" Then Match = "N'" & Request("LenMatch" & LenID(i)) & "'" Else Match = "NULL"
			sql = sql & "set @LenID = " & LenID(i) & " " & _
				"if not exists(select '' from OLKShipmentLengthMatch where ShipTypeID = @ShipTypeID and LenID = @LenID) begin " & _
				"	insert OLKShipmentLengthMatch(ShipTypeID,LenID, Match) values(@ShipTypeID, @LenID, " & Match & ") " & _
				"end else begin " & _
				"	update OLKShipmentLengthMatch set Match = " & Match & " where ShipTypeID = @ShipTypeID and LenID = @LenID " & _
				"end "
		Next
	End If

	If Request("WeiID") <> "" Then
		WeiID = Split(Request("WeiID"), ", ")
		sql = sql & "declare @WeiID int "
		For i = 0 to UBound(WeiID)
			If Request("WeiMatch" & WeiID(i)) <> "" Then Match = "N'" & Request("WeiMatch" & WeiID(i)) & "'" Else Match = "NULL"
			sql = sql & "set @WeiID = " & WeiID(i) & " " & _
				"if not exists(select '' from OLKShipmentWeightMatch where ShipTypeID = @ShipTypeID and WeiID = @WeiID) begin " & _
				"	insert OLKShipmentWeightMatch(ShipTypeID,WeiID, Match) values(@ShipTypeID, @WeiID, " & Match & ") " & _
				"end else begin " & _
				"	update OLKShipmentWeightMatch set Match = " & Match & " where ShipTypeID = @ShipTypeID and WeiID = @WeiID " & _
				"end "
		Next
	End If
	
	conn.execute(sql)
	
	conn.close
	
	If Request("btnApply") <> "" Then
		Response.Redirect "adminShipSis.asp?ShipTypeID=" & Request("ShipTypeID")
	Else
		Response.Redirect "adminShipSis.asp"
	End If
End Sub


Private Sub adminLogos()

	If Request("TopLogo") <> "" Then TopLogo = "N'" & saveHTMLDecode(Request("TopLogo"), False) & "'" Else TopLogo = "NULL"
	If Request("MailLogo") <> "" Then MailLogo = "N'" & saveHTMLDecode(Request("MailLogo"), False) & "'" Else MailLogo = "NULL"
	If Request("AgentLogo") <> "" Then AgentLogo = "N'" & saveHTMLDecode(Request("AgentLogo"), False) & "'" Else AgentLogo = "NULL"
	sql = "update OLKCommon set TopLogo = " & TopLogo & ", MailLogo = " & MailLogo & ", AgentLogo = " & AgentLogo & ", LastUpdate = getdate()"
	conn.execute(sql)
	myApp.LoadAdminLogos
	conn.close
	myApp.ResetLastUpdate
	Response.Redirect "adminLogos.asp"
End Sub

Private Sub adminPrintTitle()

	Select Case Request("cmd")
		Case "a"
			If Request("chkShowName") = "Y" Then ShowName = "Y" Else ShowName = "N"
			sql = "declare @LineIndex int set @LineIndex = ISNULL((select max(LineIndex)+1 from OLKDocAddHdr),1) " & _
				  "select @LineIndex LineIndex " & _
				  "insert OLKDocAddHdr(LineIndex, Name, Query, Access, Row, Col, ShowName) " & _
				  "values(@LineIndex,N'" & saveHTMLDecode(Request("lineName"), False) & "',N'" &  saveHTMLDecode(Request("Query"), False) & "','" & Request("Access") & "', " & Request("Row") & ", " & Request("Col") & ", '" & ShowName & "')"
			set rs = Server.CreateObject("ADODB.RecordSet")
			set rs = conn.execute(sql)
			rI = rs(0)
			rs.close
			set rs = nothing
			
			If Request("lineNameTrad") <> "" Then
				SaveNewTrad Request("lineNameTrad"), "DocAddHdr", "LineIndex", "alterName", rI
			End If
			
			If Request("QueryDef") <> "" Then
				SaveNewDef Request("QueryDef"), rI
			End If
			
			conn.close
		Case "u"
			ArrVal = Split(Request("LineIndex"),", ")
			For i = 0 to UBound(ArrVal)
				sql = sql & " update OLKDocAddHdr set Name = N'" & saveHTMLDecode(Request("lineName" & ArrVal(i)), False) & "', " & _
				"Row = " & Request("Row" & ArrVal(i)) & ", Col = " & Request("Col" & ArrVal(i)) & ", Access = '" & Request("Access" & ArrVal(i)) & "' " & _
				"where LineIndex = " & ArrVal(i)
			Next
			If sql <> "" Then conn.execute(sql)
		Case "del"
			sql = "delete OLKDocAddHdr where LineIndex = " & Request("rI") & _
					"delete OLKDocAddHdrAlterNames where LineIndex = " & Request("rI")
			conn.execute(sql)
		Case "e"
			If Request("chkShowName") = "Y" Then ShowName = "Y" Else ShowName = "N"
			sql = "update OLKDocAddHdr set Name = N'" & saveHTMLDecode(Request("lineName"), False) & "', " & _
					"Query = N'" & saveHTMLDecode(Request("Query"), False) & "', Row = " & Request("Row") & _
				  ", Col = " & Request("Col") & ", Access = '" & Request("Access") & "', ShowName = '" & ShowName & "' " & _
				  "where LineIndex = " & Request("rI")
			conn.execute(sql)
			rI = Request("rI")
	End Select
	If Request("btnApply") <> "" Then
		response.Redirect "adminPrintTitle.asp?edit=Y&rI=" & rI & "&1=1#table20"
	Else
		response.Redirect "adminPrintTitle.asp"
	End If
End Sub

Private Sub menuGroups()
	set rs = Server.CreateObject("ADODB.RecordSet")
	
	If Request("cmd") = "remGroup" Then
		sql = "declare @GroupID int set @GroupID = " & Request("GroupID") & " " & _
				"delete OLKMenuGroups where GroupID = @GroupID " & _
				"delete OLKMenuGroupsAlterNames where GroupID = @GroupID " & _
				"delete OLKMenuGroupsLines where GroupID = @GroupID " & _
				"delete OLKMenuGroupsLinesQryGroups where GroupID = @GroupID"
		conn.execute(sql)
	ElseIf Request("cmd") = "update" Then
		sql = "select GroupID from OLKMenuGroups"
		set rs = conn.execute(sql)
		sql = ""
		do while not rs.eof
			GroupID = Replace(rs("GroupID"), "-", "_")
			AllPosition = Request("AllPosition" & GroupID)
			Active = Request("chkActive" & GroupID)
			Ordr = Request("GroupOrder" & GroupID)
			If Active = "" Then Active = "N"
			
			addStr = ""
			If rs("GroupID") >= 0 Then
				addStr = ", GroupName = N'" & Request("GroupName" & GroupID) & "' "
			End If
			sql = sql & "update OLKMenuGroups set AllPosition = '" & AllPosition & "', Active = '" & Active & "', Ordr = " & Ordr & addStr & " " & _
						"where GroupID = " & rs("GroupID") & " "
		rs.movenext
		loop
		sql = sql & "update OLKCommon set DefMenuGroup = " & Request("rdDefault")
		
		conn.execute(sql)
	ElseIf Request("cmd") = "edit" and Request("btnAddLine") = "" Then
		If Request("chkActive") = "Y" Then Active = "Y" Else Active = "N"
		If Request("SearchFilter") = "" Then SearchFilter = "NULL" Else SearchFilter = "N'" & saveHTMLDecode(Request("SearchFilter"), False) & "'"
		sql = "update OLKMenuGroups set GroupName = N'" & Request("GroupName") & "', SearchFilter = " & SearchFilter & ", AllPosition = '" & Request("AllPosition") & "', Active = '" & Active & "', Ordr = " & Request("GroupOrder") & " " & _
				"where GroupID = " & Request("editID")
		conn.execute(sql)
		
		sql = "select LineID, TableType from OLKMenuGroupsLines where GroupID = " & Request("editID")
		set rs = conn.execute(sql)
		sql = ""
		do while not rs.eof
			If Request("cmbDescTable" & rs("LineID")) <> "" Then DescTable = "N'" & Request("cmbDescTable" & rs("LineID")) & "'" Else DescTable = "NULL"
			If Request("cmbDescID" & rs("LineID")) <> "" Then DescID = "N'" & Request("cmbDescID" & rs("LineID")) & "'" Else DescID = "NULL"
			If Request("cmbDescName" & rs("LineID")) <> "" Then DescName = "N'" & Request("cmbDescName" & rs("LineID")) & "'" Else DescName = "NULL"
			If Request("FilterFormula" & rs("LineID")) <> "" Then FilterFormula = "N'" & saveHTMLDecode(Request("FilterFormula" & rs("LineID")), False) & "'" Else FilterFormula = "NULL"
			If Request("DescFormula" & rs("LineID")) <> "" Then DescFormula = "N'" & saveHTMLDecode(Request("DescFormula" & rs("LineID")), False) & "'" Else DescFormula = "NULL"
			
			sql = sql & "update OLKMenuGroupsLines set DescTable = " & DescTable & ", DescID = " & DescID & ", DescName = " & DescName & ", Ordr = " & Request("LineOrder" & rs("LineID")) & ",  " & _
					"FilterFormula = " & FilterFormula & ", DescFormula = " & DescFormula & " " & _
					"where GroupID = " & Request("editID") & " and LineID = " & rs("LineID") & " " & _
					"delete OLKMenuGroupsLinesQryGroups where GroupID = " & Request("editID") & " and LineID = " & rs("LineID") & " "
			
			If rs("TableType") = "Q" Then
				sql = sql & "insert OLKMenuGroupsLinesQryGroups(GroupID, LineID, ItmsTypCod) " & _
							"select " & Request("editID") & ", " & rs("LineID") & ", Value from OLKCommon.dbo.OLKSplit('" & Request("QryGroups" & rs("LineID")) & "', ',') "
			End If
		rs.movenext
		loop
		If sql <> "" Then conn.execute(sql)
		
		groupID = Request("editID")
		
		If Request("btnApply") <> "" Then redirVal = "?editID=" & Request("editID")
	ElseIf Request("cmd") = "edit" and Request("btnAddLine") <> "" Then
		Select Case Request("cmbType")
			Case "S"
				Table = "'" & Request("cmbTableField") & "'"
				Select Case Request("cmbTableField")
					Case "OITB"
						FilterID = "'ItmsGrpCod'"
						FilterName = "'ItmsGrpNam'"
					Case "OMRC"
						FilterID = "'FirmCode'"
						FilterName = "'FirmName'"
					Case "OCRD"
						FilterID = "'CardCode'"
						FilterName = "'CardName'"
				End Select
			Case "U"
				Table = "NULL"
				FilterID = "'" & Request("cmbTableField") & "'"
				FilterName = "NULL"
			Case "F"
				Table = "NULL"
				FilterID = "'" & Request("cmbTableField") & "'"
				FilterName = "NULL"
			Case "Q"
				Table = "NULL"
				FilterID = "NULL"
				FilterName = "NULL"
			Case "T"
				Table = "'" & Request("cmbTableField") & "'"
				FilterID = "'" & Request("cmbTableFilterID") & "'"
				FilterName = "NULL"
		End Select
		If Request("cmbDescTable") <> "" Then cmbDescTable = "N'" & Request("cmbDescTable") & "'" Else cmbDescTable = "NULL"
		If Request("cmbDescID") <> "" Then cmbDescID = "N'" & Request("cmbDescID") & "'" Else cmbDescID = "NULL"
		If Request("cmbDescName") <> "" Then cmbDescName = "N'" & Request("cmbDescName") & "'" Else cmbDescName = "NULL"
		If Request("FilterFormulaAdd") <> "" Then FilterFormula = "N'" & saveHTMLDecode(Request("FilterFormulaAdd"), False) & "'" Else FilterFormula = "NULL"
		If Request("DescFormulaAdd") <> "" Then DescFormula = "N'" & saveHTMLDecode(Request("DescFormulaAdd"), False) & "'" Else DescFormula = "NULL"
		sql = "declare @GroupID int set @GroupID = " & Request("editID") & " " & _
				"declare @LineID int set @LineID = IsNull((select Max(LineID)+1 from OLKMenuGroupsLines where GroupID = @GroupID), 0) " & _
				"insert OLKMenuGroupsLines(GroupID, LineID, TableType, QueryTable, FilterID, FilterName, FilterFormula, DescTable, DescID, DescName, DescFormula, Ordr) " & _
				"values(@GroupID, @LineID, '" & Request("cmbType") & "', " & Table & ", " & FilterID & ", " & FilterName & ", " & FilterFormula & ", " & cmbDescTable & ", " & _
					cmbDescID & ", " & cmbDescName & ", " & DescFormula & ", " & Request("AddOrder") & ") "
		If Request("cmbType") = "Q" Then
			sql = sql & "insert OLKMenuGroupsLinesQryGroups(GroupID, LineID, ItmsTypCod) " & _
						"select @GroupID, @LineID, Value from OLKCommon.dbo.OLKSplit('" & Request("QryGroupsAdd") & "', ',') "
		End If
		conn.execute(sql)
		
		groupID = Request("editID")

		redirVal = "?editID=" & Request("editID")
	ElseIf Request("cmd") = "new" Then
		If Request("chkActive") = "Y" Then Active = "Y" Else Active = "N"
		If Request("SearchFilter") = "" Then SearchFilter = "NULL" Else SearchFilter = "N'" & saveHTMLDecode(Request("SearchFilter"), False) & "'"
		sql = 	"declare @GroupID int set @GroupID = IsNull((select Max(GroupID)+1 from OLKMenuGroups), 0) " & _
				"select @GroupID GropuID " & _
				"insert OLKMenuGroups(GroupID, GroupName, SearchFilter, AllPosition, Active, Ordr) " & _
				"values(@GroupID, N'" & Request("GroupName") & "', " & SearchFilter & ", '" & Request("AllPosition") & "', '" & Active & "', " & Request("GroupOrder") & ")"
		set rs = conn.execute(sql)
		
		If Request("GroupNameTrad") <> "" Then
			SaveNewTrad Request("GroupNameTrad"), "MenuGroups", "GroupID", "alterGroupName", rs(0)
		End If
		
		If Request("SearchFilterDef") <> "" Then
			SaveNewDef Request("SearchFilterDef"), CStr(DefID) & rs(0)
		End If


		If Request("btnApply") <> "" Then redirVal = "?editID=" & rs(0)
		groupID = rs(0)
	ElseIf Request("cmd") = "remLine" Then
		sql = "delete OLKMenuGroupsLines where GroupID = " & Request("GroupID") & " and LineID = " & Request("LineID") & _
		" delete OLKMenuGroupsLinesQryGroups where GroupID = " & Request("GroupID") & " and LineID = " & Request("LineID")
		conn.execute(sql)
		redirVal = "?editID=" & Request("GroupID")
		groupID = Request("GroupID")
	End If
		
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKGenMenuGroup" & Session("ID")
	cmd.Parameters.Refresh
	cmd("@GroupID") = groupID
	cmd.execute()

	conn.close
	
	Response.Redirect "adminMenuGroups.asp" & redirVal
End Sub

Private Sub adminUsersLic()
	uSql = "declare @AlertType char(1) set @AlertType = 'S' " & _
	"declare @AlertID int set @AlertID = 2 " & _
	"declare @ToType char(1) declare @ToID int " & _
	"insert OLKAlertsTo(AlertType, AlertID, ToType, ToID, SendIntrnl, SendEMail, SendSMS, SendFax, AlertLang) " & _
	"select T0.AlertType, T0.AlertID, 'O', T1.SlpCode, 'N', 'N', 'N', 'N', NULL " & _
	"from OLKAlerts T0 " & _
	"cross join OLKAgentsAccess T1 " & _
	"where not exists(select '' from OLKAlertsTo where AlertType = T0.AlertType and AlertID = T0.AlertID and ToType = 'O' and ToID = T1.SlpCode) " & _
	"insert OLKAlertsTo(AlertType, AlertID, ToType, ToID, SendIntrnl, SendEMail, SendSMS, SendFax, AlertLang) " & _
	"select T0.AlertType, T0.AlertID, 'S', T1.UserID, 'N', 'N', 'N', 'N', NULL " & _
	"from OLKAlerts T0 " & _
	"cross join OUSR T1 " & _
	"where not exists(select '' from OLKAlertsTo where AlertType = T0.AlertType and AlertID = T0.AlertID and ToType = 'S' and ToID = T1.UserID) "
	

	set rs = server.CreateObject("ADODB.RecordSet")
	sql = 	"select T1.SlpCode " & _
			"from OLKAgentsAccess T0 " & _
			"inner join OSLP T1 on T1.SlpCode = T0.SlpCode " & _
			"left outer join OLKAlertsTo T2 on T2.ToType = 'O' and T2.ToID = T1.SlpCode and T2.AlertType = 'S' and T2.AlertID = 2 " & _
			"where T0.Access <> 'D' "
	set rs = conn.execute(sql)

	uSql = uSql & " set @ToType = 'O' "
	do while not rs.eof
		If Request("cmbLngO" & rs("SlpCode")) <> "" Then Lang = "'" & Request("cmbLngO" & rs("SlpCode")) & "'" Else Lang = "NULL"
		uSql = uSql & 	"set @ToID = " & rs("SlpCode") & " " & _
						"	update OLKAlertsTo set AlertLang = " & Lang & " " & _
						"	where AlertType = @AlertType and ToType = @ToType and ToID = @ToID "
	rs.movenext
	loop
	sql = 	"select T0.USERID " & _
			"from OUSR T0 " & _
			"left outer join OLKAlertsTo T2 on T2.ToType = 'S' and T2.ToID = T0.USERID and T2.AlertType = 'S' and T2.AlertID = 2 " & _
			"where T0.Groups <> 99 "
	set rs = conn.execute(sql)

	uSql = uSql & " set @ToType = 'S' "
	do while not rs.eof
		If Request("cmbLngS" & rs("USERID")) <> "" Then Lang = "'" & Request("cmbLngS" & rs("USERID")) & "'" Else Lang = "NULL"
		uSql = uSql & 	"set @ToID = " & rs("UserID") & " " & _
						"	update OLKAlertsTo set AlertLang = " & Lang & " " & _
						"	where AlertType = @AlertType and ToType = @ToType and ToID = @ToID " & _
	rs.movenext
	loop
	
	conn.execute(uSql)
	
	setActiveMail
	
	conn.close
	Response.Redirect "adminUsersLng.asp"
End Sub

Private Sub adminDefinition

	
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.CommandType = adCmdStoredProc
	cmd.ActiveConnection = connCommon
	cmd.CommandText = "DBOLKSetQryDefinition" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@PageID") = Request("PageID")
	cmd("@FieldID") = Request("FieldID")
	cmd("@FieldKey") = Request("FieldKey")
	cmd("@Definition") = Request("txtDefinition")
	cmd.execute
	
	WinClose("")
End Sub

Private Sub adminTrad()

	
	Table = Request("Table")
	'ColumnID = Request("ColumnID")
	'ID = Request("ID")
	arrCol = Split(Request("ColumnID"), ",")
	arrVal = Split(Request("ID"), ",")

	ColumnName = Request("ColumnName")
	
	sql = "declare @LanID int "
	For j = 0 to UBound(myLanIndex)
		LanID = myLanIndex(j)(4)
		sql = sql & "set @LanID = " & LanID & " "
		If Request("txt" & LanID) <> "" Then
			sql = sql & "if not exists(select 'A' from OLK" & Table & "AlterNames where LanID = @LanID "
			
			For i = 0 to UBound(arrCol)
				If IsNumeric(arrVal(i)) Then
					sql = sql & "and " & arrCol(i) & " = " & arrVal(i) & " "
				Else
					sql = sql & "and " & arrCol(i) & " = N'" & arrVal(i) & "' "
				End If
			Next

			sql = sql & ") begin " & _
			"	insert OLK" & Table & "AlterNames(LanID, "
			
			For i = 0 to UBound(arrCol)
				sql = sql & arrCol(i) & ", "
			Next
			
			sql = sql & ColumnName & ") " & _
			"	values(@LanID, "
			
			For i = 0 to UBound(arrCol)
				If IsNumeric(arrVal(i)) Then
					sql = sql & arrVal(i) & ", "
				Else
					sql = sql & "N'" & arrVal(i) & "', "
				End If
			Next
			
			sql = sql & "N'" & saveHTMLDecode(Request("txt" & LanID), False) & "') " & _
			"end else begin " & _
			"	update OLK" & Table & "AlterNames set " & ColumnName & " = N'" & saveHTMLDecode(Request("txt" & LanID), False) & "' " & _
			"	where LanID = @LanID "
			
			For i = 0 to UBound(arrCol)
				If IsNumeric(arrVal(i)) Then
					sql = sql & "and " & arrCol(i) & " = " & arrVal(i) & " "
				Else
					sql = sql & "and " & arrCol(i) & " = N'" & arrVal(i) & "' "
				End If
			Next

			sql = sql & "end "
		Else
			sql = sql & "	update OLK" & Table & "AlterNames set " & ColumnName & " = NULL " & _
			"	where LanID = @LanID "
			
			For i = 0 to UBound(arrCol)
				If IsNumeric(arrVal(i)) Then
					sql = sql & "and " & arrCol(i) & " = " & arrVal(i) & " "
				Else
					sql = sql & "and " & arrCol(i) & " = N'" & arrVal(i) & "' "
				End If
			Next
		End If
	Next
	
	conn.execute(sql)

	WinClose("")
End Sub

Private Sub adminMyData()
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.CommandType = adCmdStoredProc
	cmd.ActiveConnection = connCommon
	cmd.CommandText = "DBOLKAdminClientSettings" & Session("ID")
	cmd.Parameters.Refresh()
	
	If Request("MyDataReadOnly") = "Y" Then MyDataReadOnly = "Y" Else MyDataReadOnly = "N"
	If Request("EnableDROnlyNote") = "Y" Then EnableDROnlyNote = "Y" Else EnableDROnlyNote = "N"
	
	cmd("@MyDataReadOnly") = MyDataReadOnly
	cmd("@EnableDROnlyNote") = EnableDROnlyNote
	
	If Request("DataReadOnlyNote") <> "" Then cmd("@DataReadOnlyNote") = Request("DataReadOnlyNote")
	
	cmd.execute()
	
	myApp.LoadClientSettings
	myApp.ResetLastUpdate
	
	Response.Redirect "adminMyData.asp"
End Sub

Private Sub adminCartMore()

	If Request("ShowCSearchTree") = "Y" Then ShowCSearchTree = "Y" Else ShowCSearchTree = "N"
	If Request("ShowCAdSearch") = "Y" Then ShowCAdSearch = "Y" Else ShowCAdSearch = "N"
	sql = "update OLKCommon set ShowCSearchTree = '" & ShowCSearchTree & "', ShowCAdSearch = '" & ShowCAdSearch & "'"
	conn.execute(sql)
	Response.Redirect "adminCartMore.asp"
End Sub

Private Sub adminNavCat()

	set cmd = Server.CreateObject("ADODB.Command")
	cmd.CommandType = adCmdStoredProc
	cmd.ActiveConnection = connCommon
	cmd.CommandText = "DBOLKAdminCatNav" & Session("ID")
	cmd.Parameters.Refresh()
	
	If Request("cmd") = "del" Then
		cmd("@NavIndex") = Request("delIndex")
		cmd("@Action") = "D"
		cmd.Execute()
		retVal = "adminCatNav.asp"
	End If
	
	If Request("cmd") = "edit" or Request("cmd") = "add" Then
		If Request("Active") <> "" Then Active = "Y" Else Active = "N"
		If Request("AutoRedir") = "Y" Then AutoRedir = "Y" Else AutoRedir = "N"
		If Request("ApplyAnonCatFilter") = "Y" Then ApplyAnonCatFilter = "Y" Else ApplyAnonCatFilter = "N"
		
		cmd("@NavTitle") = saveHTMLDecode(Request("NavTitle"), True)
		cmd("@NavType") = saveHTMLDecode(Request("NavType"), True)
		cmd("@NavImgType") = Request("NavImgType")
		cmd("@Access") = Request("Access")
		cmd("@Active") = Active
		cmd("@AutoRedir") = AutoRedir
		cmd("@ApplyAnonCatFilter") = ApplyAnonCatFilter
		
		If Request("NavDesc") <> "" Then cmd("@NavDesc") = saveHTMLDecode(Request("NavDesc"), True)
		If Request("NavImg") <> "" Then cmd("@NavImg") = Request("NavImg")
		If Request("NavImgQry") <> "" and Request("NavImgType") = "Q" Then cmd("@NavImgQry") = saveHTMLDecode(Request("NavImgQry"), True)
		If Request("NavQry") <> "" and Request("NavType") = "Q" Then cmd("@NavQry") = saveHTMLDecode(Request("NavQry"), True)
		If Request("CatType") <> "" Then cmd("@CatType") = Request("CatType")
		If Request("ShowFrom") <> "" Then cmd("@ShowFrom") = Request("ShowFrom")
		If Request("ShowTo") <> "" Then cmd("@ShowTo") = Request("ShowTo")
		
		If Request("cmd") = "edit" Then
			cmd("@NavIndex") = Request("editIndex")
			cmd("@Action") = "U"
		End If
		
		cmd.execute()
		NavIndex = cmd("@NavIndex").value
		
		If Request("NavTitleTrad") <> "" Then
			SaveNewTrad Request("NavTitleTrad"), "CatNav", "NavIndex", "AlterNavTitle", NavIndex
		End If
		
		If Request("NavDescTrad") <> "" Then
			SaveNewTrad Request("NavDescTrad"), "CatNav", "NavIndex", "AlterNavDesc", NavIndex
		End If
				
		If Request("NavQryDef") <> "" Then
			SaveNewDef Request("NavQryDef"), NavIndex
		End If
				
		If Request("NavImgQryDef") <> "" Then
			SaveNewDef Request("NavImgQryDef"), NavIndex
		End If
		
		If Request("cmd") = "edit" Then
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = adCmdStoredProc
			cmd.CommandText = "DBOLKClearData" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@TableID") = "CatNavSub"
			cmd("@Index") = Request("editIndex")
			cmd.execute()
		End If
		
		If Request("SubIndex") <> "" Then
			ArrVal = Split(Request("SubIndex"), ", ")
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = adCmdStoredProc
			cmd.CommandText = "DBOLKAdminCatNavSub" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@NavIndex") = NavIndex
			For i = 0 to UBound(ArrVal)
				cmd("@SubIndex") = ArrVal(i)
				cmd.execute()
			Next
			
		End If
		
		If Request("btnApply") <> "" Then
			retVal = "adminCatNav.asp?editIndex=" & NavIndex
		Else
			retVal = "adminCatNav.asp"
		End If
	End If
	
	If Request("cmd") = "NavIndex" Then
		sql = "update OLKCommon set NavIndexByX = " & Request("NavIndexByX") & ", NavIndexByY = " & Request("NavIndexByY") & " " & _
				"delete OLKCatNavIndex "
		
		SecCount = CInt(Request("NavIndexByX"))*CInt(Request("NavIndexByY"))
		For i = 1 to SecCount
			sql = sql & "insert OLKCatNavIndex (ID, NavIndex) values(" & i & ", " & Request("NavID" & i) & ") "
		Next
	
		conn.execute(sql)
		retVal = "adminCatNav.asp"
	End If
	
	conn.close
	Response.Redirect retVal
End Sub

Private Sub adminBN()

	
	GroupId = Request("GroupId")
	Select Case Request("cmd")
		Case "uGrp"
			sql = "select GroupID from olkBNGroups order by groupId"
			set rs = conn.execute(sql)
			sql = ""
			do while not rs.eof
				sql = sql & " update olkBNGroups set GroupName = N'" & saveHTMLDecode(Request("txtGroupName" & rs("GroupID")), False) & "', SizeX = '" & Request("txtSizeX" & rs("GroupID"))&"', SizeY = '" & Request("txtSizeY" & rs("GroupID"))  & "' where GroupID = " & rs("GroupId")
				rs.movenext
			loop
			
			If sql <> "" Then conn.execute(sql)		
			
			If Request("NewGroupName") <> "" Then
				
				strSQL = "Declare @GroupID int set @GroupID = IsNull((select Max(GroupID)+1 from OLKBNGroups),0) " & _
						  "select @GroupID GroupID " & _
						  "INSERT INTO OLKBNGroups(GroupID ,GroupName, SizeX ,SizeY) " &  _
						  "SELECT @GroupID, N'" & saveHTMLDecode(Request("NewGroupName"), False) & "', '"&Request("newSizeX")&"', '"&Request("newSizeY")&"' "		
						  
				set rs = Server.CreateObject("ADODB.RecordSet")
				set rs = conn.Execute(strSQL)
				newGroupId = rs(0)
				
				If Request("NewGroupNameTrad") <> "" Then
					SaveNewTrad Request("NewGroupNameTrad"), "BNGroups", "GroupID ", "AlterGroupName", newGroupID 
				End If
			End If
			retVal = "adminBN.asp?GroupId=" & GroupID
		Case "remGR"
			sql = "delete OLKBNGroups where GroupID = " & Request("delId") & " " & _
					"delete OLKBNGroupsAlterNames where GroupID = " & Request("delId")
			conn.execute(sql)	
			retVal = "adminBN.asp?GroupId=" & GroupID
		Case "uActive"
			strSQL = "select  BannerID from OLKBN where Status in ('A','N') "
			If GroupId <> "" Then
				strSQL = strSQL &	 "and GroupID = "&GroupID
			End If
			set rs = conn.execute(strSql)
			sql = ""
			do while not rs.eof
				If Request("chkStatus" & rs(0)) = "A" Then Status = "A" Else Status = "N"
				sql = sql & "update olkBN set Status = '" & Status & "' where BannerID = " & rs(0) & " "
				rs.movenext
			loop
			
			If sql <> "" Then conn.execute(sql)
			retVal = "adminBN.asp?GroupId=" & GroupId
		Case "remBN"
			strSQL = "update OLKBN set status='D' "& _
					  "where BannerID = " & Request("BannerID")
			conn.execute(strSQL)
			retVal = "adminBN.asp?GroupId="&GroupId
		Case "saveBN"
			BannerID = Request("BannerID")
			If Request("txtBannerDesc") <> "" Then BannerDesc = "N'" & saveHTMLDecode(Request("txtBannerDesc"), False) & "'" Else BannerDesc = "NULL"
			If Request("txtStartDate") <> "" Then StartDate = "Convert(datetime,'" & SaveSqlDate(Request("txtStartDate")) & "',120)" Else StartDate = "NULL"
			If Request("txtEndDate") <> "" Then EndDate = "Convert(datetime,'" & SaveSqlDate(Request("txtEndDate")) & "',120)" Else EndDate = "NULL"
			If Request("txtQuery") <> "" Then Query = "N'" & saveHTMLDecode(Request("txtQuery"), False) & "'" Else Query = "NULL"
			If Request("lstStatus") = "A" Then Status = "A" Else Status = "N"
			If BannerID <> "" Then
				sql = 	 "update OLKBN set " & _
						 "BannerDesc = " & BannerDesc & ", "& _
			 			 "Link = '" & Request("txtLink") & "', "& _  
						 "Picture = '" & Request("txtPicture") & "',  "& _
						 "Query = " & Query & ", "& _
						 "GroupID = " & Request("lstGroupID") & ", "& _
						 "StartDate = " & StartDate & ", "& _
						 "EndDate = " & EndDate & ", "& _
						 "Status = '" & Status & "' "& _
						 "Where BannerID = "& BannerID
				conn.Execute(sql)
			Else
				sql = 	"declare @BannerID int set @BannerID = IsNull((select Max(BannerID)+1 from OLKBN),0) " & _
						"select @BannerID BannerID " & _
						"insert OLKBN(BannerID, BannerDesc, Link, Picture, Query, GroupID, StartDate, EndDate, Status) " & _
						"values(@BannerID, " & BannerDesc & ", '" & Request("txtLink") & "', '" & Request("txtPicture") & "', " & _
						Query & ", " & Request("lstGroupID") & ", " & StartDate & ", " & EndDate & ", '" & Status & "') "
				set rs = conn.execute(sql)
				BannerID = rs("BannerID")
				

				If Request("BannerDescTrad") <> "" Then
					SaveNewTrad Request("BannerDescTrad"), "BN", "BannerID", "AlterBannerDesc", BannerID
				End If
				
				If Request("BannerDescDef") <> "" Then
					SaveNewDef Request("BannerDescDef"), BannerID
				End If
			End If

			
			sql = 	"declare @BannerID int set @BannerID = " & BannerID & " " & _
					"delete OLKBNOCRG where BannerID = @BannerID "
			If Request("hdnCodClientes") <> "" Then
				arr = Split(Request("hdnCodClientes"), ", ")
				For i = 0 to UBound(arr)
					sql = sql & "insert OLKBNOCRG(BannerID, GroupCode) values(@BannerID, " & arr(i) & ") "
				Next
			End If
			
			sql = sql & "delete OLKBNOCRY where BannerID = @BannerID "
			If Request("hdnCodPaises") <> "" Then
				arr = Split(Request("hdnCodPaises"), ", ")
				For i = 0 to UBound(arr)
					sql = sql & "insert OLKBNOCRY(BannerID, CountryID) values(@BannerID, " & arr(i) & ") "
				Next
			End If
			
			sql = sql & "delete OLKBNSections where BannerID = @BannerID "
			If Request("hdnCodSecciones") <> "" Then
				arr = Split(Request("hdnCodSecciones"), ", ")
				For i = 0 to UBound(arr)
					sql = sql & "insert OLKBNSections(BannerID, SecType, SecID) values(@BannerID, '" & Mid(arr(i),2,1) & "', " & Mid(arr(i), 3, Len(arr(i))-3) & ") "
				Next
			End If
			
			conn.execute(sql)
			
			If Request("btnApply") <> "" Then
				retVal = "adminBNEdit.asp?BannerID=" & BannerID & "&GroupId=" & Request("GroupId")
			Else
				retVal = "adminBN.asp?GroupId=" & Request("GroupId")
			End If
	End Select
	conn.close
	Response.Redirect retVal
End Sub


Private Sub adminSecIndex()

	
	sql = "update OLKCommon set SecIndexByX = " & Request("SecIndexByX") & ", SecIndexByY = " & Request("SecIndexByY") & " " & _
			"delete OLKSecIndex "
	
	SecCount = CInt(Request("SecIndexByX"))*CInt(Request("SecIndexByY"))
	For i = 1 to SecCount
		sql = sql & "insert OLKSecIndex(ID, SecID) values(" & i & ", " & Request("SecID" & i) & ") "
	Next

	conn.execute(sql)
	conn.close
	
	If Request("btnApply") <> "" Then
		Response.Redirect "adminSecIndex.asp"
	Else
		Response.Redirect "adminSec.asp?UType=C"
	End If
End Sub

Private Sub adminObjs()

	set rs = Server.CreateObject("ADODB.RecordSet")
	Select Case Request("uCmd")
		Case "update"
			sql = "select ObjType, ObjID from OLKObjects where Status <> 'D'"
			set rs = conn.execute(sql)
			sql = ""
			do while not rs.eof
				If Request("Status" & rs("ObjType") & rs("ObjID")) = "Y" Then Status = "Y" Else Status = "N"
				sql = sql & "update OLKObjects set Status = '" & Status & "' "
				sql = sql & " where ObjType = N'" & rs("ObjType") & "' and ObjId = " & rs("ObjId") & " "
			rs.movenext
			loop
			If sql <> "" Then conn.execute(sql)
			Response.Redirect "adminDefObjs.asp"
		Case "edit"
			If Request("Status") = "Y" Then Status = "Y" Else Status = "N"
			If Request("btnRestore") = "" Then
				ObjContent = "N'" & Replace(Request("ObjContent"), "'", "''") & "'"
			Else
				ObjContent = " T1.ObjContent "
			End If
			
			sql = "update OLKObjects set ObjContent = " & ObjContent & ", Status = '" & Status & "' "
			
			sql = sql & " from OLKObjects T0 "
			
			If Request("btnRestore") <> "" Then
				sql = sql & " inner join OLKCommon..OLKObjects T1 on T1.ObjType = T0.ObjType collate database_default and T1.ObjId = T0.ObjID "
			End If
			
			sql = sql & " where T0.ObjType = N'" & Request("ObjType") & "' and T0.ObjId = " & Request("ObjId")
			conn.execute(sql)
			
			sql = "select VarId, VarType from OLKObjectsVars where ObjType = '" & Request("ObjType") & "' and ObjId = " & Request("ObjId")
			set rs = conn.execute(sql)
			sql = ""
			do while not rs.eof
				sql = sql & "update OLKObjectsVars set VarValue = "
				Select Case rs("VarType")
					Case "A"
						sql = sql & "N'" & Request("VarVal" & rs("VarId")) & "' "
					Case "N"
						sql = sql & Request("VarVal" & rs("VarId"))
				End Select
				sql = sql & " where ObjType = '" & Request("ObjType") & "' and ObjId = " & Request("ObjId") & " and VarId = " & rs("VarId") & " "
			rs.movenext
			loop
			If sql <> "" Then conn.execute(sql)
			
			If CInt(Request("ObjId")) = 6 Then 
				set cmd = Server.CreateObject("ADODB.Command")
				cmd.ActiveConnection = connCommon
				cmd.CommandType = &H0004
				cmd.CommandText = "DBOLKGenSpecQry" & Session("ID")
				cmd.Parameters.Refresh
				cmd("@Name") = "OLKPromotions"
				cmd.execute()
			End If
			If Request("btnApply") <> "" or Request("btnRestore") <> "" Then
				Response.Redirect "adminDefObjEdit.asp?ObjType=" & Request("ObjType") & "&ObjId=" & Request("ObjId")
			Else
				Response.Redirect "adminDefObjs.asp"
			End If
	End Select
End Sub

Private Sub adminCatProp()

	
	If Request("ShowClientRef") = "Y" Then ShowClientRef = "Y" Else ShowClientRef = "N"
	If Request("CatShowProm") = "Y" Then CatShowProm = "Y" Else CatShowProm = "N"
	If Request("ShowClientSalUn") = "Y" Then ShowClientSalUn = "Y" Else ShowClientSalUn = "N"
	If Request("ShowPocketImg") = "Y" Then ShowPocketImg = "Y" Else ShowPocketImg = "N"
	If Request("ShowClientImg") = "Y" Then ShowClientImg = "Y" Else ShowClientImg = "N"
	If Request("ShowAgentImg") = "Y" Then ShowAgentImg = "Y" Else ShowAgentImg = "N"
	If Request("UnEmbPriceSet") = "Y" Then UnEmbPriceSet = "Y" Else UnEmbPriceSet = "N"
	If Request("EnableOfertToDisc") = "Y" Then EnableOfertToDisc = "Y" Else EnableOfertToDisc = "N"
	If Request("ShowNotAvlInv") = "Y" Then ShowNotAvlInv = "Y" Else ShowNotAvlInv = "N"
	If Request("ShowSearchTreeCount") = "Y" Then ShowSearchTreeCount = "Y" Else ShowSearchTreeCount = "N"
	If Request("ShowSearchTreeSubCount") = "Y" Then ShowSearchTreeSubCount = "Y" Else ShowSearchTreeSubCount = "N"
	If Request("EnableSearchAlterCode") = "Y" Then EnableSearchAlterCode = "Y" Else EnableSearchAlterCode = "N"
	If Request("ShowQtyInUnAg") = "Y" Then ShowQtyInUnAg = "Y" Else ShowQtyInUnAg = "N"
	If Request("ShowQtyInUnCl") = "Y" Then ShowQtyInUnCl = "Y" Else ShowQtyInUnCl = "N"
	If Request("SearchExactA") = "Y" Then SearchExactA = "Y" Else SearchExactA = "N"
	If Request("SearchExactC") = "Y" Then SearchExactC = "Y" Else SearchExactC = "N"
	If Request("SearchExactP") = "Y" Then SearchExactP = "Y" Else SearchExactP = "N"
	If Request("ShowPriceTax") = "Y" Then ShowPriceTax = "Y" Else ShowPriceTax = "N"
	If Request("SearchByVendorCode") = "Y" Then SearchByVendorCode = "Y" Else SearchByVendorCode = "N"
	If Request("EnableUnitSelection") = "Y" Then EnableUnitSelection = "Y" Else EnableUnitSelection = "N"
	If Request("EnableMultCheck") = "Y" Then EnableMultCheck = "Y" Else EnableMultCheck = "N"
	
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = adCmdStoredProc
	cmd.CommandText = "DBOLKAdminCatProp" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@CarArt") 				= Request("CarArt")
	cmd("@AutoSearchOpen")		= Request("AutoSearchOpen")
	cmd("@olkItemReport2")		= Request("olkItemReport2")
	cmd("@UnEmbPriceSet")		= UnEmbPriceSet 
	cmd("@AgentSaleUnit")		= Request("AgentSaleUnit")
	cmd("@ClientSaleUnit")		= Request("ClientSaleUnit")
	cmd("@CatShowProm")			= CatShowProm 
	cmd("@ShowClientRef")		= ShowClientRef 
	cmd("@DefCatOrdrC")			= Request("DefCatOrdrC")
	cmd("@DefCatOrdrV")			= Request("DefCatOrdrV")
	cmd("@ShowClientSalUn")		= ShowClientSalUn 
	cmd("@ShowPocketImg")		= ShowPocketImg  
	cmd("@ShowClientImg")		= ShowClientImg
	cmd("@ShowAgentImg")		= ShowAgentImg
	cmd("@EnableOfertToDisc")	= EnableOfertToDisc
	cmd("@ShowNotAvlInv")		= ShowNotAvlInv
	cmd("@ShowSearchTreeCount") = ShowSearchTreeCount
	cmd("@ShowSearchTreeSubCount") = ShowSearchTreeSubCount
	cmd("@EnableSearchAlterCode") = EnableSearchAlterCode
	cmd("@DefViewCL")				= Request("DefViewCL")
	cmd("@DefViewAG")				= Request("DefViewAG")
	cmd("@ShowQtyInUnAg")			= ShowQtyInUnAg
	cmd("@ShowQtyInUnCl")			= ShowQtyInUnCl
	cmd("@SearchExactA")			= SearchExactA
	cmd("@SearchMethodA")			= Request("SearchMethodA")
	cmd("@SearchExactC")			= SearchExactC
	cmd("@SearchMethodC")			= Request("SearchMethodC")
	cmd("@SearchExactP")			= SearchExactP
	cmd("@SearchMethodP")			= Request("SearchMethodP")
	cmd("@ShowPriceTax")			= ShowPriceTax
	cmd("@SearchByVendorCode")		= SearchByVendorCode
	cmd("@EnableUnitSelection")		= EnableUnitSelection
	cmd("@EnableMultCheck")			= EnableMultCheck
	
	cmd.execute()
	
	sql = "delete OLKSearchQryGroups "
	
	If Request("chkQryGroup") <> "" Then
		sql = sql & "insert OLKSearchQryGroups(ItmsTypCod) select Value from OLKCommon.dbo.OLKSplit('" & Request("chkQryGroup") & "', ', ') "
	End If
	
	conn.execute(sql)
	
	myApp.LoadAdminCatProp
	myApp.ResetLastUpdate
	
	conn.close
	response.redirect "adminCatProp.asp"
End Sub

Private Sub adminCart()

	
	If Request("EnableCartSum") = "Y" Then EnableCartSum = "Y" Else EnableCartSum = "N"
	If Request("EnTop10Items") = "Y" Then EnTop10Items = "Y" Else EnTop10Items = "N"
	If Request("ExpItems") = "Y" Then ExpItems = "Y" Else ExpItems = "N"
	If Request("BasketMItems") = "Y" Then BasketMItems = "Y" Else BasketMItems = "N"
	If Request("SDKLineMemo") = "Y" Then SDKLineMemo = "Y" Else SDKLineMemo = "N"
	If Request("EnCSelDoc") = "Y" Then EnCSelDoc = "Y" Else EnCSelDoc = "N"
	If Request("PrintCCartNote") = "Y" Then PrintCCartNote = "Y" Else PrintCCartNote = "N"
	If Request("EnableCartImpC") = "Y" Then EnableCartImpC = "Y" Else EnableCartImpC = "N"
	If Request("EnableCartImpV") = "Y" Then EnableCartImpV = "Y" Else EnableCartImpV = "N"
	If Request("EnableClientMDoc") = "Y" Then EnableClientMDoc = "Y" Else EnableClientMDoc = "N"
	If Request("EnableDiscount") = "Y" Then EnableDiscount = "Y" Else EnableDiscount = "N"
	If Request("ShowPriceBefDiscount") = "Y" Then ShowPriceBefDiscount = "Y" Else ShowPriceBefDiscount = "N"
	If Request("ApplyMaxDiscToSU") = "Y" Then ApplyMaxDiscToSU = "Y" Else ApplyMaxDiscToSU = "N"
	If Request("UseCustomTransMsg") = "Y" Then UseCustomTransMsg = "Y" Else UseCustomTransMsg = "N"
	If Request("ShowLineDiscount") = "Y" Then ShowLineDiscount = "Y" Else ShowLineDiscount = "N"
	If Request("PrintPriceBefDiscount") = "Y" Then PrintPriceBefDiscount = "Y" Else PrintPriceBefDiscount = "N"
	If Request("PrintLineDiscount") = "Y" Then PrintLineDiscount = "Y" Else PrintLineDiscount = "N"
	If Request("AllowClientPartSuppSel") = "Y" Then AllowClientPartSuppSel = "Y" Else AllowClientPartSuppSel = "N"
	If Request("EnSelAll") = "Y" Then EnSelAll = "Y" Else EnSelAll = "N"
	If Request("EnableHideCartHdr") = "Y" Then EnableHideCartHdr = "Y" Else EnableHideCartHdr = "N"
	If Request("EnableDocPrjSel") = "Y" Then EnableDocPrjSel = "Y" Else EnableDocPrjSel = "N"
	If Request("EnableAnonCart") = "Y" Then EnableAnonCart = "Y" Else EnableAnonCart = "N"
	If Request("FastAddUnRem") = "Y" Then FastAddUnRem = "Y" Else FastAddUnRem = "N"
	If Request("FastAddBeep") = "Y" Then FastAddBeep = "Y" Else FastAddBeep = "N"
	If Request("EnableMultiBPCart") = "Y" Then EnableMultiBPCart = "Y" Else EnableMultiBPCart = "N"
	If Request("CartItmBarCode") = "Y" Then CartItmBarCode = "Y" Else CartItmBarCode = "N"
	
				
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandText = "DBOLKAdminCart" & Session("ID")
	cmd.CommandType = adCmdStoredProc
	cmd.Parameters.Refresh()
	cmd("@AfterCartAddC") 			= Request("AfterCartAddC")
	cmd("@AfterCartAddV") 			= Request("AfterCartAddV")
	cmd("@AfterCartAddPocket") 		= Request("AfterCartAddPocket")
	cmd("@D_DocC") 					= Request("D_DocC")
	cmd("@CartSumQty") 				= Request("CartSumQty")
	cmd("@EnableCartSum") 			= EnableCartSum
	cmd("@Top10Items") 				= Request("Top10Items")
	cmd("@EnTop10Items") 			= EnTop10Items
	cmd("@ExpItems") 				= ExpItems
	cmd("@BasketMItems") 			= BasketMItems
	cmd("@CCartNote") 				= saveHTMLDecode(Request("CCartNote"), True)
	cmd("@DocMCBal") 				= Request("DocMCBal")
	cmd("@CartGroup") 				= Request("CartGroup")
	cmd("@SDKLineMemo") 			= SDKLineMemo
	cmd("@EnCSelDoc") 				= EnCSelDoc
	cmd("@PrintCCartNote") 			= PrintCCartNote
	cmd("@PrintCCartNote") 			= PrintCCartNote
	cmd("@PocketDefDoc")			= Request("PocketDefDoc")
	cmd("@EnableCartImpC")			= EnableCartImpC
	cmd("@EnableCartImpV")			= EnableCartImpV
	cmd("@EnableClientMDoc")		= EnableClientMDoc
	cmd("@EnableDiscount")			= EnableDiscount
	cmd("@ShowPriceBefDiscount")	= ShowPriceBefDiscount
	cmd("@ApplyMaxDiscToSU")		= ApplyMaxDiscToSU
	cmd("@MaxDiscount")				= CDbl(getNumericOut(Request("MaxDiscount")))
	cmd("@CartType")				= Request("CartType")
	cmd("@UseCustomTransMsg")		= UseCustomTransMsg
	If Request("CustomTransMsg") <> "" Then cmd("@CustomTransMsg").value = Request("CustomTransMsg")
	cmd("@ShowLineDiscount")		= ShowLineDiscount
	cmd("@PrintPriceBefDiscount")	= PrintPriceBefDiscount 
	cmd("@PrintLineDiscount")		= PrintLineDiscount 
	cmd("@AllowClientPartSuppSel")	= AllowClientPartSuppSel
	cmd("@EnSelAll")				= EnSelAll
	cmd("@EnSellAllUnitFrom")		= Request("EnSellAllUnitFrom")
	cmd("@EnableHideCartHdr")		= EnableHideCartHdr
	cmd("@EnableDocPrjSel")			= EnableDocPrjSel
	cmd("@EnableAnonCart") 			= EnableAnonCart
	If Request("AnonCartClient") <> "" Then cmd("@AnonCartClient") = Request("AnonCartClient")
	cmd("@FastAddUnRem")			= FastAddUnRem
	cmd("@FastAddBeep")				= FastAddBeep
	If Request("ItemDescModQry") <> "" Then cmd("@ItemDescModQry") = Request("ItemDescModQry")
	cmd("@EnableMultiBPCart")		= EnableMultiBPCart
	cmd("@CartItmBarCode")			= CartItmBarCode
	cmd.execute()
	
	setActiveMail()
	
	myApp.LoadAdminCart
	myApp.ResetLastUpdate

	conn.close
	response.redirect "adminCart.asp"

End Sub

Private Sub alterNames()
	set rs = Server.CreateObject("ADODB.RecordSet")
	
	sql = "select AlterID from OLKAlterNames"
	set rs = conn.execute(sql)
	sql = ""
	
	do while not rs.eof
		If Request("Sing" & rs("AlterID")) <> "" Then Sing = "N'" & saveHTMLDecode(Request("Sing" & rs("AlterID")), False) & "'" Else Sing = "NULL"
		If Request("Plur" & rs("AlterID")) <> "" Then Plur = "N'" & saveHTMLDecode(Request("Plur" & rs("AlterID")), False) & "'" Else Plur = "NULL"
		sql = sql & "update OLKAlterNames set Singular = " & Sing & ", Plural = " & Plur & " where LanID = " & Request("AlterLng") & " and AlterID = " & rs("AlterID") & " "
	rs.movenext
	loop
	conn.execute(sql)
	
	conn.close
	response.Redirect "adminAlterNames.asp?AlterLng=" & Request("AlterLng")
End Sub

Private Sub adminMsg()
	cmd = Request("cmd")

	
	If Request("Query") <> "" Then Query = "N'" & saveHTMLDecode(Request("Query"), False) & "'" Else Query = "NULL"
	sql = "update OLKAutMsg set Query = " & Query & ", Header = N'" & saveHTMLDecode(Request("Header"), False) & "', Message = N'" & saveHTMLDecode(Request("Message"), False) & "' " & _
	"where MsgID = " & Request("MsgID")
	conn.execute(sql)
	
	conn.close
	If Request("btnSave") <> "" Then
		response.Redirect "adminMsg.asp"
	Else
		response.Redirect "adminMsg.asp?MsgID=" & Request("MsgID")
	End If
End Sub

Private Sub adminAcctRejReasons()
	cmd = Request("cmd")

	set rs = Server.CreateObject("ADODB.RecordSet")	
	
	Select Case cmd
		Case "d"
			sql = "delete OLKAcctRejectNotes where ReasonIndex = " & Request("rIndex")
		Case "e"
			sql = "update OLKAcctRejectNotes set ReasonName = N'" & saveHTMLDecode(Request("ReasonName"), False) & "', Reason = N'" & saveHTMLDecode(Request("Reason"), False) & "' " & _
				  "where ReasonIndex = " & Request("rIndex")
		Case "a"
			sql = "declare @ReasonIndex int set @ReasonIndex = IsNull((select Max(ReasonIndex)+1 from OLKAcctRejectNotes), 0) " & _
				  "insert OLKAcctRejectNotes(ReasonIndex, ReasonName, Reason) values(@ReasonIndex, N'" & saveHTMLDecode(Request("ReasonName"), False) & "', " & _
				  "N'" & saveHTMLDecode(Request("Reason"), False) & "') "
	End Select
	
	conn.execute(sql)
	conn.close
	If cmd = "d" Then
		response.redirect "adminAnonLogin.asp"
	Else
		WinClose("adminAnonLogin.asp")
	End If
End Sub

Private Sub adminSec()
	RedirVal = "adminSec.asp?UType=" & Request("UType")

	set rs = Server.CreateObject("ADODB.RecordSet")
	
	Select Case Request("uCmd")
		Case "update"
			sql = "select SecType, SecID from OLKSections where Status <> 'D' and UserType = '" & Request("UType") & "'"
			set rs = conn.execute(sql)
			sql = ""
			do while not rs.eof
				rAdd = rs("SecType") & rs("SecID")
				If Request("Status" & rAdd) = "Y" Then Status = "A" Else Status = "N"
				If Request("HideMainMenu" & rAdd) = "Y" Then HideMainMenu = "Y" Else HideMainMenu = "N"
				If Request("HideSecondMenu" & rAdd) = "Y" Then HideSecondMenu = "Y" Else HideSecondMenu = "N"
				
				sql = sql & "update OLKSections set SecOrder = " & Request("SecOrder" & rAdd) & " "
				Select Case rs("SecType") 
					Case "U" 
						If Request("ReqLogin" & rAdd) = "Y" or Request("Type" & rAdd) = "R" and Request("UType") = "C" Then Login = "Y" Else Login = "N"
						sql = sql & ", SecName = N'" & saveHTMLDecode(Request("SecName" & rAdd), False) & "', " & _
									"ReqLogin = '" & Login & "', HideMainMenu = '" & HideMainMenu & "', HideSecondMenu = '" & HideSecondMenu & "'"
					Case "S"
						If rs("SecID") = 0 or rs("SecID") = 3 Then sql = sql & ", HideMainMenu = '" & HideMainMenu & "'"
						If rs("SecID") = 3 Then sql = sql & ", HideSecondMenu = '" & HideSecondMenu & "'"
						If rs("SecID") = 0 Then 
							If Request("ReqLogin" & rAdd) = "Y" or Request("Type" & rAdd) = "R" and Request("UType") = "C" Then Login = "Y" Else Login = "N"
							sql = sql & ", ReqLogin = '" & Login & "'"
						End If
				End Select
				sql = sql & ", Status = '" & Status & "'  where SecType = '" & rs("SecType") & "' and SecID = " & rs("SecID") & " "
			rs.movenext
			loop
		Case "del"
			sql = "update OLKSections set Status = 'D' where SecType = 'U' and SecID = " & Request("SecID")
		Case "edit"
			If Request("NewReqLogin") = "Y" or Request("Type") = "R" and Request("UType") = "C" Then ReqLogin = "Y" Else ReqLogin = "N"
			If Request("NewActive") = "Y" Then Status = "A" Else Status = "N"
			If Request("HideMainMenu") = "Y" Then HideMainMenu = "Y" Else HideMainMenu = "N"
			If Request("HideSecondMenu") = "Y" Then HideSecondMenu = "Y" Else HideSecondMenu = "N"
			If Request("NewManual") = "Y" Then Manual = "Y" Else Manual = "N"
			If Request("ApplyCSS") = "Y" Then ApplyCSS = "Y" Else ApplyCSS = "N"
			If Request("Form") = "Y" Then Form = "Y" Else Form = "N"
			If Request("FormScript") <> "" Then FormScript = "'" & saveHTMLDecode(Request("FormScript"), False) & "'" Else FormScript = "NULL"
			If Request("FormConfirmContent") <> "" Then FormConfirmContent =  "'" & Replace(saveHTMLDecode(Request("FormConfirmContent"), False), strURL, "") & "'" Else FormConfirmContent = "NULL"
			If Request("FormQry") <> "" Then FormQry = "'" & saveHTMLDecode(Request("FormQry"), False) & "'" Else FormQry = "NULL"
			If Request("SecContentEnableQry") = "Y" Then SecContentEnableQry = "Y" Else SecContentEnableQry = "N"
			If Request("SecContentQry") <> "" Then SecContentQry = "N'" & saveHTMLDecode(Request("SecContentQry"), False) & "'" Else SecContentQry = "NULL"
			If Request("FormQryRS") = "Y" Then FormQryRS = "Y" Else FormQryRS = "N"
			If Request("FormQryLoop") <> "" Then FormQryLoop = "N'" & saveHTMLDecode(Request("FormQryLoop"), False) & "'" Else FormQryLoop = "NULL"
			
			strURL = GetHTTPStr & Request.ServerVariables("SERVER_NAME") & Replace(Request.ServerVariables("URL"), "admin/adminSubmit.asp", "")
			Select Case Request("Type")
				Case "N"
					strContent = Replace(saveHTMLDecode(Request("NewContent"), False), strURL, "")
				Case "L"
					strContent = saveHTMLDecode(Request("NewLink"), False)
				Case "R"
					strContent = Request("rsIndex")
			End Select
			strSmallText = Replace(saveHTMLDecode(Request("SecSmallText"), False), strURL, "")
			
			If Request("SecID") = "" Then
				sql = "declare @SecID int set @SecID = IsNull((select Max(SecID)+1 from OLKSections where SecType = 'U'), 0) " & _
				"select @SecID SecID " & _
				"insert OLKSections (SecType, SecID, SecName, SecContent, SecSmallText, SecOrder, Form, Manual, ApplyCSS, HideMainMenu, HideSecondMenu, ReqLogin, UserType, Type, Status, CreateDate, SecContentEnableQry, SecContentQry, FormQryRS, FormQryLoop) " & _
				"values('U', @SecID, N'" & saveHTMLDecode(Request("NewName"), False) & "', N'" & strContent & "', N'" & strSmallText & "', " & Request("NewOrder") & ", '" & Form & "', N'" & Manual & "', '" & ApplyCSS & "', " & _
				"'" & HideMainMenu & "', '" & HideSecondMenu & "', '" & ReqLogin & "', '" & Request("UType") & "', '" & Request("Type") & "', '" & Status & "', getdate(), '" & SecContentEnableQry & "', " & SecContentQry & ", '" & FormQryRS & "', " & FormQryLoop & ") "
				set rs = conn.execute(sql)
				SecID = rs(0)
				sql = ""
				
				If Request("SecNameTrad") <> "" Then
					SaveNewTrad Request("SecNameTrad"), "Sections", "SecType,SecID", "AlterSecName", "U, " & SecID
				End If
				If Request("SecContentTrad") <> "" Then
					SaveNewTrad Request("SecContentTrad"), "Sections", "SecType,SecID", "AlterSecContent", "U, " & SecID
				End If
				If Request("SecSmallTextTrad") <> "" Then
					SaveNewTrad Request("SecSmallTextTrad"), "Sections", "SecType,SecID", "AlterSecSmallText", "U, " & SecID
				End If
			Else
				SecID = Request("SecID")
				sql = "update OLKSections set SecName = N'" & saveHTMLDecode(Request("NewName"), False) & "', SecContent = N'" & strContent & "', SecSmallText = N'" & strSmallText & "', " & _
				"SecOrder = " & Request("NewOrder") & ", Form = '" & Form & "', FormScript = " & FormScript & ", FormConfirmContent = " & FormConfirmContent & ", FormQry = " & FormQry & ", Manual = N'" & Manual & "', ApplyCSS = '" & ApplyCSS & "', HideMainMenu = N'" & HideMainMenu & "', HideSecondMenu = '" & HideSecondMenu & "', ReqLogin = N'" & ReqLogin & "', Type = '" & Request("Type") & "', Status = N'" & Status & "', " & _
				"SecContentEnableQry = '" & SecContentEnableQry & "', SecContentQry = " & SecContentQry & ", FormQryRS = '" & FormQryRS & "', FormQryLoop = " & FormQryLoop & " " &_
				"where SecType = 'U' and SecID = " & Request("SecID")
			End If
			If Request("btnApply") <> "" Then RedirVal = "adminSecEdit.asp?SecID=" & SecID & "&rCount=" & Request("rCount") & "&UType=" & Request("UType")
	End Select
	if sql <> "" Then conn.execute(sql)
	conn.close 
	response.redirect RedirVal
End Sub 

Private Sub updateNews()

	      
    If Request("newsSource") <> "" Then newsSource = "N'" & saveHTMLDecode(Request("newsSource"), False) & "'" else newsSource = "NULL"
    If Request("newsImg") <> "" Then newsImg = "N'" & Request("newsImg") & "'" Else newsImg = "NULL"
    If Request("chkActive") = "A" Then chkActive = "A" Else chkActive = "N"
    
	varx = Request.ServerVariables("URL")

	varText = saveHTMLDecode(Request("newsText"), False)
	varText = Replace(varText,GetHTTPStr & Request.ServerVariables("SERVER_NAME") & Mid(varx,1,Len(varx)-16),"")

sql = "update olkNews set newsTitle = N'" & saveHTMLDecode(Request("newsTitle"), False) & "', newsDate = Convert(datetime,'" & SaveSqlDate(Request("newsDate")) & "',120), " & _
	  "newsSmallText = N'" & saveHTMLDecode(Request("newsSmallText"), False) & "', newsText = N'" & varText & "', newsSource = " & newsSource & ", " & _
	  "newsImg = " & newsImg & ", Status = '" & chkActive & "' where newsIndex = " & Request("newsIndex")
	conn.execute(sql)
	conn.close 
	If Request("btnSave") <> "" Then
		response.redirect "adminNews.asp"
	Else
		response.redirect "adminNewsEdit.asp?newsIndex=" & Request("newsIndex")
	End If
End Sub

Private Sub addNews()
	set rs = Server.CreateObject("ADODB.RecordSet")
	
	sql = "select ISNULL(max(newsIndex)+1,1) newsIndex from olkNews"
	set rs = conn.execute(sql)
	newsIndex = rs("newsIndex")
	set rs = nothing
	
    If Request("newsSource") <> "" Then newsSource = "N'" & saveHTMLDecode(Request("newsSource"), False) & "'" else newsSource = "NULL"
    If Request("newsImg") <> "" Then newsImg = "N'" & Request("newsImg") & "'" Else newsImg = "NULL"
    If Request("chkActive") = "A" Then chkActive = "A" Else chkActive = "N"

	varx = Request.ServerVariables("URL")
	varText = saveHTMLDecode(Request("newsText"), False)
	varText = Replace(varText,GetHTTPStr & Request.ServerVariables("SERVER_NAME") & Mid(varx,1,Len(varx)-21),"")

sql = "insert olkNews(newsIndex, newsTitle, newsDate, newsSmallText, newsText, newsSource, newsImg, Status) " & _
	  "values(" & newsIndex & ", N'" & saveHTMLDecode(Request("newsTitle"), False) & "', Convert(datetime,'" & SaveSqlDate(Request("newsDate")) & "',120), " & _
	  "N'" & saveHTMLDecode(Request("newsSmallText"), False) & "', N'" & varText & "', " & newsSource & ",  " & newsImg & ", '" & chkActive & "')"
	conn.execute(sql)
	
	If Request("newsTitleTrad") <> "" Then
		SaveNewTrad Request("newsTitleTrad"), "News", "newsIndex", "alterNewsTitle", newsIndex
	End If
	
	If Request("newsSourceTrad") <> "" Then
		SaveNewTrad Request("newsSourceTrad"), "News", "newsIndex", "alterNewsSource", newsIndex
	End If

	If Request("newsSmallTextTrad") <> "" Then
		SaveNewTrad Request("newsSmallTextTrad"), "News", "newsIndex", "alterNewsSmallText", newsIndex
	End If
	
	If Request("newsTextTrad") <> "" Then
		SaveNewTrad Request("newsTextTrad"), "News", "newsIndex", "alterNewsText", newsIndex
	End If
	
	conn.close 
	
	If Request("btnSave") <> "" Then
		response.redirect "adminNews.asp"
	Else
		response.redirect "adminNewsEdit.asp?newsIndex=" & newsIndex
	End If
End Sub

Private Sub adminAnonLogin()
	set rs = Server.CreateObject("ADODB.RecordSet")
	If Request("EnableAnSesion") = "Y" Then EnableAnSesion = "Y" Else EnableAnSesion = "N"
	If Request("EnableAnReg") = "Y" Then EnableAnReg = "Y" Else EnableAnReg = "N"
	If Request("RegActMailAdd") <> "" Then RegActMailAdd = "N'" & Replace(Request("RegActMailAdd"), "'", "''") & "'" Else RegActMailAdd = "NULL"
	If Request("RemPwdMailAdd") <> "" Then RemPwdMailAdd = "N'" & Replace(Request("RemPwdMailAdd"), "'", "''") & "'" Else RemPwdMailAdd = "NULL"
	If Request("AnSesListNum") <> "" Then AnSesListNum = Request("AnSesListNum") Else AnSesListNum = "NULL"
	If Request("EnableAnRegTerms") = "Y" Then EnableAnRegTerms = "Y" Else EnableAnRegTerms = "N"
	If Request("AnTerms") <> "" Then AnTerms = "N'" & saveHTMLDecode(Request("AnTerms"), False) & "'" Else AnTerms = "NULL"
	if Request("AnRegConfAsignSLP") = "Y" Then AnRegConfAsignSLP = "Y" Else AnRegConfAsignSLP = "N"
	If Request("AnRegConfRejNote") = "Y" Then AnRegConfRejNote = "Y" Else AnRegConfRejNote = "N"
	If Request("EnChooseCType") = "Y" Then EnChooseCType = "Y" Else EnChooseCType = "N"
	If Request("AnRegConfFrom") <> "" Then AnRegConfFrom = Request("AnRegConfFrom") Else AnRegConfFrom = "NULL"
	If Request("AnRegConfTo") <> "" Then AnRegConfTo = Request("AnRegConfTo") Else AnRegConfTo = "NULL"
	If Request("AnonSesFilter") <> "" Then AnonSesFilter = "N'" & saveHTMLDecode(Request("AnonSesFilter"), False) & "'" Else AnonSesFilter = "NULL"
	
	If Request("WebAddress") <> "" Then
		set rs = Server.CreateObject("ADODB.RecordSet")
		sql = "EXEC OLKCommon..OLKValidateDomain N'" & Replace(Request("WebAddress"), "'", "''") & "', N'" & Session("olkdb") & "'"
		set rs = conn.execute(sql)
		If Not rs.Eof Then
			If rs(0) = "Y" Then
				WebAddress = "WebAddress"
				WebErr = "&WebAddErr=True&Address=" & Request("WebAddress") & "&usedBy=" & rs(2)
			Else
				WebAddress = "'" & Replace(Request("WebAddress"), "'", "''") & "'"
			End If
		Else
			WebAddress = "'" & Replace(Request("WebAddress"), "'", "''") & "'"
		End If
	Else
		WebAddress = "NULL"
	End If
	
	sql = "update OLKCommon set EnableAnSesion = '" & EnableAnSesion & "', EnableAnReg = '" & EnableAnReg & "', AnRegAct = '" & Request("AnRegAct") & "', " & _
			"WebAddress = " & WebAddress & ", RegActMailAdd = " & RegActMailAdd & ", RemPwdMailAdd = " & RemPwdMailAdd & ", AnSesListNum = " & AnSesListNum & ", " & _
			"EnableAnRegTerms = '" & EnableAnRegTerms & "', AnTerms = " & AnTerms & ", " & _
			"AnRegConfAsignSLP = '" & AnRegConfAsignSLP & "', AnRegConfRejNote = '" & AnRegConfRejNote & "', ClientType = '" & Request("ClientType") & "', " &  _
			"EnChooseCType = '" & EnChooseCType & "', AnRegConfFrom = " & AnRegConfFrom & ", AnRegConfTo = " & AnRegConfTo & ", AnonSesFilter = " & AnonSesFilter & ", LastUpdate = getdate()"
	conn.execute(sql)

	setActiveMail()
	
	myApp.LoadAdminAnonLogin
	myApp.ResetLastUpdate

	conn.close
	response.redirect "adminAnonLogin.asp?u=1" & WebErr

End Sub

Private Sub adminDocFlow()
	set rs = Server.CreateObject("ADODB.RecordSet")
	
	If Request("btnRestore") = "" Then
	
		If Request("FlowActive") = "Y" Then FlowActive = "Y" Else FlowActive = "N"
		If Request("LineQuery") <> "" Then LineQuery = "N'" & saveHTMLDecode(Request("LineQuery"), False) & "'" Else LineQuery = "NULL"
		If Request("NoteBuilder") = "Y" Then NoteBuilder = "Y" Else NoteBuilder = "N"
		If Request("NoteQuery") <> "" Then NoteQuery = "N'" & saveHTMLDecode(Request("NoteQuery"), False) & "'" Else NoteQuery = "NULL"
		If Request("ApplyToClient") = "Y" Then ApplyToClient = "Y" Else ApplyToClient = "N"
		If Request("FlowDraft") = "Y" Then Draft = "Y" Else Draft = "N"
		If Request("FlowAuthorize") = "Y" Then Authorize = "Y" Else Authorize = "N"
		If Request("ExecAt") <> "" Then ExecAt = "N'" & Request("ExecAt") & "'" Else ExecAt = "ExecAt"
		If Request("FlowID") <> "" Then
			sql = "declare @FlowID int set @FlowID = " & Request("FlowID") & " " & _
			"delete OLKUAF1 where FlowID = @FlowID delete OLKUAF2 where FlowID = @FlowID delete OLKUAF4 where FlowID = @FlowID " & _
			"update OLKUAF set Name = N'" & saveHTMLDecode(Request("FlowName"), False) & "', Type = " & Request("FlowType") & ", [Order] = " & Request("Order") & ", ApplyToClient = '" & ApplyToClient & "', ExecAt = " & ExecAt & ", " & _
			"Query = N'" & saveHTMLDecode(Request("FlowQuery"), False) & "', LineQuery = " & LineQuery & ", NoteBuilder = N'" & NoteBuilder & "', NoteQuery = " & NoteQuery & ", " & _
			"NoteText = N'" & saveHTMLDecode(Request("NoteText"), False) & "', Active = '" & FlowActive & "', Draft = '" & Draft & "', Authorize = '" & Authorize & "' where FlowID = @FlowID "
		Else
			sql = "declare @FlowID int set @FlowID = IsNull((select Max(FlowID)+1 from OLKUAF),0) select @FlowID NewID " & _
			"insert OLKUAF(FlowID, Name, Type, [Order], ExecAt, ApplyToClient, Query, LineQuery, NoteBuilder, NoteQuery, NoteText, Active, Draft, Authorize) " & _
			"values(@FlowID, N'" & saveHTMLDecode(Request("FlowName"), False) & "', " & Request("FlowType") & ", " & Request("Order") & ", '" & Request("ExecAt") & "', '" & ApplyToClient & "', " & _
			"N'" & saveHTMLDecode(Request("FlowQuery"), False) & "', " & LineQuery & ", N'" & NoteBuilder & "', " & NoteQuery & ", " & _
			"N'" & saveHTMLDecode(Request("NoteText"), False) & "', '" & FlowActive & "', '" & Draft & "', '" & Authorize & "') "
		End If
		SlpCode = Split(Request("SlpCode"), ", ")
		For i = 0 to UBound(SlpCode)
			sql = sql & "insert OLKUAF1(FlowID, SlpCode) values(@FlowID, " & SlpCode(i) & ") "
		next
		ObjectCode = Split(Request("ObjectCode"), ", ")
		For i = 0 to UBound(ObjectCode)
			sql = sql & "insert OLKUAF2(FlowID, ObjectCode) values(@FlowID, " & ObjectCode(i) & ") "
		next
		
		arrGrp = Split(Request("GrpID"), ", ")
		For i = 0 to UBound(arrGrp)
			If Request("AsignedSLP" & arrGrp(i)) = "Y" Then AsignedSLP = "Y" Else AsignedSLP = "N"
			If Request("GrpQuery" & arrGrp(i)) = "Y" and Request("GrpValue" & arrGrp(i) & "Query") <> "" Then
				strQuery = "N'" & saveHTMLDecode(Request("GrpValue" & arrGrp(i) & "Query"), False) & "'"
			Else
				strQuery = "NULL"
			End If
			sql = sql & "insert OLKUAF4(FlowID, AutGrpID, AsignedSLP, Query, Ordr) values(@FlowID, " & arrGrp(i) & ", '" & AsignedSLP & "', " & strQuery & ", " & Request("Order" & arrGrp(i)) & ") "
		Next
		
		If Request("FlowID") <> "" Then
			conn.execute(sql)
			FlowID = Request("FlowID")
		Else
			set rs = conn.execute(sql)
			FlowID = rs(0)
			
			If Request("FlowNameTrad") <> "" Then
				SaveNewTrad Request("FlowNameTrad"), "UAF", "FlowID", "AlterName", FlowID
			End If
			
			If Request("NoteTextTrad") <> "" Then
				SaveNewTrad Request("NoteTextTrad"), "UAF", "FlowID", "AlterNoteText", FlowID
			End If
			
			If Request("FlowQueryDef") <> "" Then
				SaveNewDef Request("FlowQueryDef"), FlowID
			End If
			
			If Request("NoteQueryDef") <> "" Then
				SaveNewDef Request("NoteQueryDef"), FlowID
			End If

			If Request("LineQueryDef") <> "" Then
				SaveNewDef Request("LineQueryDef"), FlowID
			End If
		End If
	ElseIf Request("FlowID") < 0 Then
		sql = 	"declare @FlowID int set @FlowID = " & Request("FlowID") & " " & _
				"declare @Active char(1) set @Active = (select Active from OLKUAF where FlowID = @FlowID) " & _
				"delete OLKUAF where FlowID = @FlowID " & _
				"delete OLKUAF2 where FlowID = @FlowID " & _
				"insert OLKUAF select * from OLKCommon..OLKUAF where FlowID = @FlowID " & _
				"insert OLKUAF2 select * from OLKCommon..OLKUAF2 where FlowID = @FlowID " & _
				"update OLKUAF set Active = @Active where FlowID = @FlowID "
		conn.execute(sql)
		FlowID = Request("FlowID")
	End If
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKGenQry" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@Type") = "DF"
	cmd("@ID") = FlowID
	cmd.execute()
	

	conn.close
	If Request("btnSave") <> "" Then
		response.redirect "adminDocFlowUpdate.aspx?dbName=" & Session("olkdb") & "&ID=" & Session("ID")
	Else
		response.redirect "adminDocFlowUpdate.aspx?dbName=" & Session("olkdb") & "&ID=" & Session("ID") & "&FlowID=" & FlowID
	End If
End Sub

Private Sub adminNote()

	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = adCmdStoredProc
	cmd.CommandText = "DBOLKAdminSetNote" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@type") = "NPL"
	cmd("@cmd") = UCase(CStr(Request("cmd")))
	If Request("cmd") = "e" or Request("cmd") = "a" Then
		cmd("@Name") = saveHTMLDecode(Request("NoteName"), True)
		cmd("@Note") = saveHTMLDecode(Request("Note"), True)
	End If
	If Request("cmd") = "e" or Request("cmd") = "d" Then
		cmd("@Index") = Request("NoteIndex")
	End If
	cmd.execute()
	If Request("redir") <> "" Then Response.redirect "adminCart.asp" Else WinClose("adminCart.asp")
End Sub

Private Sub adminAlert()
	AlertID = CInt(Request("AlertID"))
	
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKAdminSaveAlertData" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@AlertID") = AlertID
	If Request("Status") = "A" Then cmd("@Status") = "A"
	If Request("asigned") = "Y" Then cmd("@asigned") = "Y"
	cmd.execute()
	
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKAdminGetAlertSaveList" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@ToType") = "O"
	set rs = server.CreateObject("ADODB.RecordSet")
	set rs = cmd.execute()

	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKAdminSaveAlertToData" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@AlertID") = AlertID
	cmd("@ToType") = "O"
	do while not rs.eof
		If Request("chkIntrnlO" & rs("SlpCode")) = "Y" Then SendIntrnl = "Y" Else SendIntrnl = "N"
		If Request("chkMailO" & rs("SlpCode")) = "Y" Then SendEMail = "Y" Else SendEMail = "N"
		If Request("chkRegO" & rs("SlpCode")) = "Y" Then Regular = "Y" Else Regular = "N"
		If Request("chkDraftO" & rs("SlpCode")) = "Y" Then Draft = "Y" Else Draft = "N"
		
		cmd("@ToID") = rs("SlpCode")
		cmd("@SendIntrnl") = SendIntrnl
		cmd("@SendEMail") = SendEMail
		cmd("@AlertReg") = Regular 
		cmd("@AlertDraft") = Draft 
		cmd("@BranchIndex") = Request("branchO" & rs("SlpCode"))
		cmd.execute()
	rs.movenext
	loop
	
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKAdminGetAlertSaveList" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@ToType") = "S"
	set rs = cmd.execute()

	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKAdminSaveAlertToData" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@AlertID") = AlertID
	cmd("@ToType") = "S"
	do while not rs.eof
		If Request("chkIntrnlS" & rs("USERID")) = "Y" Then SendIntrnl = "Y" Else SendIntrnl = "N"
		If Request("chkMailS" & rs("USERID")) = "Y" Then SendEMail = "Y" Else SendEMail = "N"
		If Request("chkSMSS" & rs("USERID")) = "Y" Then SendSMS = "Y" Else SendSMS = "N"
		If Request("chkFaxS" & rs("USERID")) = "Y" Then SendFax = "Y" Else SendFax = "N"
		If Request("chkRegS" & rs("USERID")) = "Y" Then Regular = "Y" Else Regular = "N"
		If Request("chkDraftS" & rs("USERID")) = "Y" Then Draft = "Y" Else Draft = "N"
		
		cmd("@ToID") = rs("USERID")
		cmd("@SendIntrnl") = SendIntrnl
		cmd("@SendEMail") = SendEMail
		cmd("@SendSMS") = SendSMS
		cmd("@SendFax") = SendFax
		cmd("@AlertReg") = Regular 
		cmd("@AlertDraft") = Draft 
		cmd("@BranchIndex") = Request("branchS" & rs("USERID"))
		cmd.execute()
	rs.movenext
	loop
	
	conn.close
	Response.Redirect "adminAlerts.asp?AlertID=" & Request("AlertID")
End Sub

Private Sub adminInv()

	If Request("ManageItmWhs") = "Y" Then ManageItmWhs = "Y" Else ManageItmWhs = "N"
	If Request("GenFAppV") = "Y" Then GenFAppV = "Y" Else GenFAppV = "N"
	If Request("GenFAppC") = "Y" Then GenFAppC = "Y" Else GenFAppC = "N"
	If Request("EnableMinInv") = "Y" Then EnableMinInv = "Y" Else EnableMinInv = "N"
	If Request("EnableMinInvV") = "Y" Then EnableMinInvV = "Y" Else EnableMinInvV = "N"
	If Request("EnableCodeBarsQry") = "Y" Then EnableCodeBarsQry = "Y" Else EnableCodeBarsQry = "N"
	If Request("VerfyDispWhs") = "S" Then VerfyDispWhs = "S" Else VerfyDispWhs = "D"
	If Request("EnableItemRec") = "Y" Then EnableItemRec = "Y" Else EnableItemRec = "N"
	If Request("chkEnableCombos") = "Y" Then EnableCombos = "Y" Else EnableCombos = "N"
	If Request("EnableSearchItmSupp") = "Y" Then EnableSearchItmSupp = "Y" Else EnableSearchItmSupp = "N"
	
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandText = "DBOLKAdminInv" & Session("ID")
	cmd.CommandType = adCmdStoredProc
	
	cmd("@MinInv") 					= CDbl(Request("MinInv"))
	cmd("@MinInvBy") 				= Request("MinInvBy")
	cmd("@MinInvV") 				= CDbl(Request("MinInvV"))
	cmd("@MinInvVBy") 				= Request("MinInvVBy")
	cmd("@MinPrice") 				= CDbl(Request("MinPrice"))
	cmd("@VerfyDisp") 				= Request("VerfyDisp")
	cmd("@VerfyDispWhs") 			= VerfyDispWhs
	cmd("@WhsCode") 				= Request("WhsCode")
	cmd("@InvBDGBy") 				= Request("InvBDGBy")
	cmd("@f_creacion") 				= Request("f_creacion")
	cmd("@ManageItmWhs") 			= ManageItmWhs
	cmd("@GenFilter") 				= Request("GenFilter")
	cmd("@GenFAppV") 				= GenFAppV
	cmd("@GenFAppC") 				= GenFAppC
	cmd("@VerfyDispMethod") 		= "C" 'Request("VerfyDispMethod")
	cmd("@EnableMinInv")			= EnableMinInv
	cmd("@EnableMinInvV")			= EnableMinInvV
	cmd("@EnableCodeBarsQry")		= EnableCodeBarsQry
	cmd("@CodeBarsQryMethod")		= Request("CodeBarsQryMethod")
	cmd("@ApplyInvFiltersBy")		= Request("ApplyInvFiltersBy")
	cmd("@EnableItemRec")			= EnableItemRec
	cmd("@EnableCombos")			= EnableCombos
	cmd("@EnableSearchItmSupp")		= EnableSearchItmSupp

	If Request("CodeBarsQry") <> "" Then cmd("@CodeBarsQry")	= Request("CodeBarsQry")
	If Request("ItemRecQry") <> "" Then cmd("@ItemRecQry") = Request("ItemRecQry")
	
	cmd.execute()
	
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKGenQry" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@Type") = "ItemRec"
	cmd.execute()
	
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKRestoreUAF" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@ExecAt") = "D2"
	
	set rs = Server.CreateObject("ADODB.RecordSet")
	sql = "select ObjectCode, Type from OLKInOutSettings"
	set rs = conn.execute(sql)
	
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandText = "DBOLKSetInOutSettings" & Session("ID")
	cmd.CommandType = adCmdStoredProc
	cmd.Parameters.Refresh()
	
	do while not rs.eof
		myID = rs("ObjectCode") & rs("Type")
		
		If Request("ChkShowReqSum" & myID) = "Y" Then ChkShowReqSum = "Y" Else ChkShowReqSum = "N"
		If Request("ChkAllowOverl" & myID) = "Y" Then ChkAllowOverload = "Y" Else ChkAllowOverload = "N"
		If Request("ChkImpExp" & myID) = "Y" Then ChkImpExp = "Y" Else ChkImpExp = "N"
		
		cmd("@ObjectCode") = rs("ObjectCode")
		cmd("@Type") = rs("Type")
		cmd("@ChkShowReqSum") = ChkShowReqSum
		cmd("@ChkAllowOverload") = ChkAllowOverload
		cmd("@ChkImpExp") = ChkImpExp
		cmd("@ChkSerial") = Request("ChkSerial" & myID)
		cmd("@ChkOp") = Request("ChkOp" & myID)
		If Request("Series" & myID) <> "" Then cmd("@ChkOpSeries") = Request("Series" & myID)
		
		cmd.execute()
	rs.movenext
	loop
	
	myApp.LoadAdminInv
	myApp.ResetLastUpdate
	
	conn.close
	response.redirect "adminInv.asp"
End Sub

Private Sub adminCUFD()
	TableID = Request.Form("TableID")
	
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.CommandType = &H0004
	cmd.ActiveConnection = connCommon
	cmd.CommandText = "DBOLKAdminCUFD" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@TableID") = TableID
	
	ArrVal = Split(Request("FieldID"),", ")
	For i = 0 to UBound(ArrVal)
		FieldID = ArrVal(i)
		
		
		cmd("@FieldID") = FieldID
		
		saveID = Replace(FieldID, "-", "_")
		AType = Request("AType" & saveID)
		
		If Request("Null" & saveID) = "Y" Then NullField = "Y" Else NullField = "N"
		If Request("Active" & saveID) = "Y" Then Active = "Y" Else Active = "N"
		If Request("Query" & saveID) = "Y" Then Query = "Y" Else Query = "N"
		cmd("@GroupID") = Request("cmbGroup" & saveID)
		cmd("@NullField") = NullField
		cmd("@Active") = Active
		If AType <> "" Then cmd("@AType") = AType Else cmd("@AType") = "T"
		If Request("OP" & saveID) <> "" Then cmd("@OP") = Request("OP" & saveID) Else cmd("@OP") = "T"
		cmd("@Query") = Query
		If CInt(FieldID) >= 0 Then 
			cmd("@Order") = Request("UDFOrder" & saveID) 
			cmd("@Pos") = Request("Pos" & saveID)
		Else 
			cmd("@Order") = -1
			cmd("@Pos") = "D"
		End If
		If Request("SqlQuery" & saveID) <> "" Then cmd("@SqlQuery") = Request("SqlQuery" & saveID) Else cmd("@SqlQuery") = Null
		If Request("SqlQueryField" & saveID) <> "" Then cmd("@SqlQueryField") = Request("SqlQueryField" & saveID) Else cmd("@SqlQueryField") = Null
		cmd.execute()
	Next
	
	GenMyQuery TableID
	
	If TableID = "CRD1" Then
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKGenQry" & Session("ID")
		cmd.Parameters.Refresh
		cmd("@Type") = "Country"
		cmd.execute()
	End If
End Sub

'Query Generators
Private Sub GenMyQuery(ByVal TableID)
	myApp.ConnectDB
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKGetMyQuery" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@TableID") = TableID
	cmd.execute()
End Sub

Private Sub adminCUFDGroups()
	sql = "declare @TableID nvarchar(20) set @TableID = '" & Request("TableID") & "' " 
	Select Case Request("cmd")
		Case "uGrp"
			If Request("GroupID") <> "" Then
				ArrVal = Split(Request("GroupID"),", ")
				For i = 0 to UBound(ArrVal)
					If CInt(ArrVal(i)) >= 0 Then GroupID = Arrval(i) Else GroupID = "_1"
					sql = sql & "update OLKCUFDGroups set GroupName = N'" & saveHTMLDecode(Request("GroupName" & GroupID), False) & "', [Order] = " & Request("GroupOrder" & GroupID) & " where TableID = @TableID and GroupID = " & ArrVal(i) & " "
				Next
				conn.execute(sql)
			End If
			
			If Request("NewGroupName") <> "" Then
				sql = "declare @TableID nvarchar(20) set @TableID = '" & Request("TableID") & "' " 
				sql = "declare @TableID nvarchar(20) set @TableID = '" & Request("TableID") & "' " & _
				"declare @GroupID int set @GroupID = (select Max(GroupID)+1 from OLKCUFDGroups where TableID = @TableID) " & _
				"select @GroupID GroupID " & _
				"insert OLKCUFDGroups(TableID, GroupID, GroupName, [Order]) " & _
				"values(@TableID, @GroupID, N'" & saveHTMLDecode(Request("NewGroupName"), False) & "', " & Request("GroupOrder") & ") "
				
				set rs = Server.CreateObject("ADODB.RecordSet")
				set rs = conn.execute(sql)
				If Request("GroupNameTrad") <> "" Then
					SaveNewTrad Request("GroupNameTrad"), "CUFDGroups", "TableID,GroupID", "AlterGroupName", Request("TableID") & "," & rs("GroupID")
				End If
			End If
		Case "delGrp"
			sql = sql & "delete OLKCUFDGroups where TableID = @TableID and GroupID = " & Request("id") & " " & _
			"delete OLKCUFDGroupsAlterNames where TableID = @TableID and GroupID = " & Request("id")
			conn.execute(sql)
	End Select
	
	conn.close
End Sub

Private Sub adminPriceCod()
	sql = " update OLKMyCod set NewKey = N'" & saveHTMLDecode(Request("Field0"), False) & "' where Type = N'" & Request("codType") & "' and OrgKey = '0'" & _
		  " update OLKMyCod set NewKey = N'" & saveHTMLDecode(Request("Field1"), False) & "' where Type = N'" & Request("codType") & "' and OrgKey = '1'" & _
		  " update OLKMyCod set NewKey = N'" & saveHTMLDecode(Request("Field2"), False) & "' where Type = N'" & Request("codType") & "' and OrgKey = '2'" & _
		  " update OLKMyCod set NewKey = N'" & saveHTMLDecode(Request("Field3"), False) & "' where Type = N'" & Request("codType") & "' and OrgKey = '3'" & _
		  " update OLKMyCod set NewKey = N'" & saveHTMLDecode(Request("Field4"), False) & "' where Type = N'" & Request("codType") & "' and OrgKey = '4'" & _
		  " update OLKMyCod set NewKey = N'" & saveHTMLDecode(Request("Field5"), False) & "' where Type = N'" & Request("codType") & "' and OrgKey = '5'" & _
		  " update OLKMyCod set NewKey = N'" & saveHTMLDecode(Request("Field6"), False) & "' where Type = N'" & Request("codType") & "' and OrgKey = '6'" & _
		  " update OLKMyCod set NewKey = N'" & saveHTMLDecode(Request("Field7"), False) & "' where Type = N'" & Request("codType") & "' and OrgKey = '7'" & _
		  " update OLKMyCod set NewKey = N'" & saveHTMLDecode(Request("Field8"), False) & "' where Type = N'" & Request("codType") & "' and OrgKey = '8'" & _
		  " update OLKMyCod set NewKey = N'" & saveHTMLDecode(Request("Field9"), False) & "' where Type = N'" & Request("codType") & "' and OrgKey = '9'" & _
		  " update OLKMyCod set NewKey = N'" & saveHTMLDecode(Request("Field_"), False) & "' where Type = N'" & Request("codType") & "' and OrgKey = '-'" & _
		  " update OLKMyCod set NewKey = N'" & saveHTMLDecode(Request("Field_11"), False) & "' where Type = N'" & Request("codType") & "' and OrgKey = '.'"
	conn.execute(sql)
	conn.close
End Sub

Private Sub adminNew()
	If Request("AsignedSLP") = "Y" Then AsignedSLP = "Y" Else AsignedSLP = "N"
	If Request("showCxcOpenInv") = "Y" Then showCxcOpenInv = "Y" Else showCxcOpenInv = "N"
	If Request("showCxcOpenInvC") = "Y" Then showCxcOpenInvC = "Y" Else showCxcOpenInvC = "N"
	If Request("showCxcDueDate") = "Y" or Request("showCxcOpenInvBy") = "DocDueDate" Then showCxcDueDate = "Y" Else showCxcDueDate = "N"
	If Request("showCxcDueDateC") = "Y" or Request("showCxcOpenInvByC") = "DocDueDate" Then showCxcDueDateC = "Y" Else showCxcDueDateC = "N"
	If Request("EnableBranchs") = "Y" Then EnableBranchs = "Y" Else EnableBranchs = "N"
	If Request("CopyLastFCRate") = "Y" Then CopyLastFCRate = "Y" Else CopyLastFCRate = "N"
	If Request("EnRetroPoll") = "Y" Then EnRetroPoll = "Y" Else EnRetroPoll = "N"
	If Request("EnBlockRClk") = "Y" Then EnBlockRClk = "Y" Else EnBlockRClk = "N"
	If Request("showCxcIncTrans") = "Y" Then showCxcIncTrans = "Y" Else showCxcIncTrans = "N"
	If Request("showCxcIncTransC") = "Y" Then showCxcIncTransC = "Y" Else showCxcIncTransC = "N"
	If Request("EnableCLogLogin") = "Y" Then EnableCLogLogin = "Y" Else EnableCLogLogin = "N"
	If Request("EnableVLogLogin") = "Y" Then EnableVLogLogin = "Y" Else EnableVLogLogin = "N"
	If Request("EnableCSearchFilterLog") = "Y" Then EnableCSearchFilterLog = "Y" Else EnableCSearchFilterLog = "N"
	If Request("EnableCSearchItemLog") = "Y" Then EnableCSearchItemLog = "Y" Else EnableCSearchItemLog = "N"
	If Request("EnableCItemViewLog") = "Y" Then EnableCItemViewLog = "Y" Else EnableCItemViewLog = "N"
	If Request("EnableCItemPurLog") = "Y" Then EnableCItemPurLog = "Y" Else  EnableCItemPurLog = "N"
	If Request("EnableCSearchByVatId") = "Y" Then EnableCSearchByVatId = "Y" Else EnableCSearchByVatId = "N"
	If Request("EnableCSearchByLicTradNum") = "Y" Then EnableCSearchByLicTradNum = "Y" Else EnableCSearchByLicTradNum = "N"
	If Request("AllowAgentAccessCDoc") = "Y" Then AllowAgentAccessCDoc = "Y" Else AllowAgentAccessCDoc = "N"
	If Request("AddUserApr") = "Y" Then AddUserApr = "Y" Else AddUserApr = "N"
	If Request("EnableAtt") = "Y" Then EnableAtt = "Y" Else EnableAtt = "N"
	If Request("ShowBPCountry") = "Y" Then ShowBPCountry = "Y" Else ShowBPCountry = "N"
	If Request("FlowAutDirectAdd") = "Y" Then FlowAutDirectAdd = "Y" Else FlowAutDirectAdd = "N"
	If Request("EnableDelegate") = "Y" Then EnableDelegate = "Y" Else EnableDelegate = "N"

	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandText = "DBOLKAdminGeneral" & Session("ID")
	cmd.CommandType = adCmdStoredProc
	cmd.Parameters.Refresh()
	cmd("@CopyLastFCRate") 				= CopyLastFCRate
	cmd("@EnBlockRClk") 				= EnBlockRClk
	cmd("@EnBlkRClkMsg") 				= saveHTMLDecode(Request("EnBlkRClkMsg"), True)
	cmd("@SlpCode") 					= Request("SlpCode")
	cmd("@AddDocSlpCode") 				= Request("AddDocSlpCode")
	cmd("@AsignedSLP") 					= AsignedSLP
	cmd("@ECDays") 						= Request("ECDays")
	cmd("@showCxcOpenInvBy") 			= Request("showCxcOpenInvBy")
	cmd("@showCxcOpenInvByC") 			= Request("showCxcOpenInvByC")
	cmd("@showCxcOpenInv") 				= showCxcOpenInv
	cmd("@showCxcOpenInvC") 			= showCxcOpenInvC
	cmd("@showCxcIncTrans") 			= showCxcIncTrans
	cmd("@showCxcIncTransC") 			= showCxcIncTransC
	cmd("@showCxcDueDate")				= showCxcDueDate
	cmd("@showCxcDueDateC")				= showCxcDueDateC
	cmd("@EnableBranchs") 				= EnableBranchs
	cmd("@EnRetroPoll") 				= EnRetroPoll
	cmd("@EnRetroPollDays") 			= Request("EnRetroPollDays")
	cmd("@EnableCLogLogin")				= EnableCLogLogin
	cmd("@EnableVLogLogin") 			= EnableVLogLogin
	cmd("@EnableCSearchFilterLog") 		= EnableCSearchFilterLog
	cmd("@EnableCSearchItemLog") 		= EnableCSearchItemLog
	cmd("@EnableCItemViewLog") 			= EnableCItemViewLog
	cmd("@EnableCItemPurLog") 			= EnableCItemPurLog
	cmd("@EnableCSearchByVatId") 		= EnableCSearchByVatId
	cmd("@EnableCSearchByLicTradNum") 	= EnableCSearchByLicTradNum
	cmd("@AllowAgentAccessCDoc") 		= AllowAgentAccessCDoc
	If Request("AgentClientsFilter") <> "" Then cmd("@AgentClientsFilter") = Request("AgentClientsFilter")
	cmd("@DefClientOPTab")				= Request("DefClientOPTab")
	cmd("@AlterLocation")				= Request("AlterLocation")
	cmd("@ActOrdr1")					= Request("ActOrdr1")
	cmd("@ActOrdr2")					= Request("ActOrdr2")
	If Request("ViewDocFilter") <> "" Then cmd("@ViewDocFilter") = Request("ViewDocFilter")
	cmd("@AddUserApr")					= AddUserApr
	cmd("@EnableAtt") 					= EnableAtt
	cmd("@ShowBPCountry") 				= ShowBPCountry
	cmd("@FlowAutDirectAdd") 			= FlowAutDirectAdd
	cmd("@EnableDelegate") 				= EnableDelegate
	cmd.execute()
	myApp.LoadAdminGeneral
	myApp.ResetLastUpdate
	conn.close
End Sub

Private Sub adminAPwd

	set oLic = server.CreateObject("TM.LicenceConnect.LicenceConnection")
	oLic.LicenceServer = licip
	oLic.LicencePort = licport

	myApp.ConnectCommon
	set rs = server.createobject("ADODB.RecordSet")
	sql = "select pwd from olkadminlogin"
	set rs = conn.execute(sql)
	If CStr(rs("pwd")) = oLic.GetEncPwd(CStr(saveHTMLDecode(Request("OldPwd"), False))) Then
		conn.execute("update olkadminlogin set pwd = N'" & oLic.GetEncPwd(saveHTMLDecode(Request("NewPwd"), False)) & "'")
		conn.close
		Response.Redirect "adminPwd.asp?ErrMsg=False"
	Else
		conn.close
		Response.Redirect "adminPwd.asp?ErrMsg=True"
	End If
End Sub

Private Sub adminPortal()
set rs = server.createobject("ADODB.RecordSet")
sql = "select ObjectCode from olkDocConf where ObjectCode in (13,17,15,23,24)"
set rs = conn.execute(sql)
sql = ""
do while not rs.eof
If Request("Confirm" & rs("ObjectCode")) = "Y" Then Confirm = "Y" Else COnfirm = "N"
sql = sql & " update olkDocConf set Confirm = '" & Confirm & "' where ObjectCode = " & rs("ObjectCode")
rs.movenext
loop
conn.execute(sql)
conn.close
set rs = nothing
response.redirect "admin.asp?cmd=portal"
End Sub

Private Sub adminBatchOpt()
Select Case Request("cmd")
	Case "a"
		rowField = saveHTMLDecode(Request("customSql"), False)
		If Request("rowTypeRnd") = "ON" Then rowTypeRnd = "Y" Else rowTypeRnd = "N"
		If Request("rowActive") = "ON" Then rowActive = "Y" Else rowActive = "N"
		sql = "declare @rowIndex int set @rowIndex = ISNULL((select max(rowIndex)+1 from olkBatchRep),1) " & _
			  "select @rowIndex rowIndex " & _
			  "insert olkBatchRep(rowIndex, rowName, rowField, rowType, rowTypeRnd, rowTypeDec, rowActive, rowOrder) " & _
			  "values(@rowIndex,N'" & saveHTMLDecode(Request("rowName"), False) & "',N'" & rowField & "','" & Request("rowType") & "', '" & rowTypeRnd & "', '" & Request("rowTypeDec") & "', '" & rowActive & "', " & Request("RowOrder") & ")"
		set rs = Server.CreateObject("ADODB.RecordSet")
		set rs = conn.execute(sql)
		rI = rs(0)
		rs.close
		set rs = nothing
		
		If Request("rowNameTrad") <> "" Then
			SaveNewTrad Request("rowNameTrad"), "BatchRep", "rowIndex", "alterRowName", rI
		End If
			
		If Request("customSqlDef") <> "" Then
			SaveNewDef Request("customSqlDef"), rI
		End If
		
		conn.close
	Case "u"
		ArrVal = Split(Request("rowIndex"),", ")
		For i = 0 to UBound(ArrVal)
			If Request("rowTypeRnd" & ArrVal(i)) = "ON" Then rowTypeRnd = "Y" Else rowTypeRnd = "N"
			If Request("rowActive" & ArrVal(i)) = "ON" Then rowActive = "Y" Else rowActive = "N"
			sql = sql & " update olkBatchRep set rowName = N'" & saveHTMLDecode(Request("rowName" & ArrVal(i)), False) & "', " & _
			"rowType = '" & Request("rowType" & ArrVal(i)) & "', rowTypeRnd = '" & rowTypeRnd & "', " & _
			"rowActive = '" & rowActive & "', rowOrder = " & Request("RowOrder" & ArrVal(i)) & " where rowIndex = " & ArrVal(i)
		Next
		If sql <> "" Then conn.execute(sql)
	Case "del"
		sql = "delete olkBatchRep where rowIndex = " & Request("rI") & _
				"delete OLKBatchRepAlterNames where rowIndex = " & Request("rI")
		conn.execute(sql)
	Case "e"
		If Request("rowTypeRnd") = "ON" Then rowTypeRnd = "Y" Else rowTypeRnd = "N"
		If Request("rowActive") = "ON" Then rowActive = "Y" Else rowActive = "N"
		sql = "update olkBatchRep set rowName = N'" & saveHTMLDecode(Request("rowName"), False) & "', " & _
				"rowField = N'" & saveHTMLDecode(Request("customSql"), False) & "', rowType = '" & Request("rowType") & _
			  "', rowTypeRnd = '" & rowTypeRnd & "', rowTypeDec = '" & Request("rowTypeDec") & "', rowActive = '" & rowActive & "', RowOrder = " & Request("RowOrder") & " where rowIndex = " & Request("rI")
		conn.execute(sql)
		rI = Request("rI")
End Select

GenMyQuery "OIBT"

If Request("btnApply") <> "" Then
	response.Redirect "adminBatchOpt.asp?edit=Y&rI=" & rI & "&1=1#table20"
Else
	response.Redirect "adminBatchOpt.asp"
End If
End Sub

Private Sub adminCartOpt()
	Select Case Request("cmd")
		Case "a"
			Field = saveHTMLDecode(Request("customSql"), False)
			If Request("TypeRnd") = "ON" Then TypeRnd = "Y" Else TypeRnd = "N"
			If Request("linkActive") = "Y" Then linkActive = "Y" Else linkActive = "N"
			If Request("linkObject") <> "" Then linkObject = Request("linkObject") Else linkObject = "NULL"
			If Request("chkDynamic") = "Y" Then Dynamic = "Y" Else Dynamic = "N"
			If Request("Align") <> "" Then Align = "'" & Request("Align") & "'" Else Align = "NULL"
			sql = "declare @ID int set @ID = ISNULL((select max(ID)+1 from olkCartRep),1) " & _
				  "select @ID ID " & _
				  "insert olkCartRep(ID, Name, Field, Access, [Type], TypeRnd, TypeDec, OP, [Order], linkActive, linkObject, Align, [Dynamic]) " & _
				  "values(@ID,N'" & saveHTMLDecode(Request("Name"), False) & "',N'" & Field & "',N'" & Request("Access") & _
				  "','" & Request("Type") & "', '" & TypeRnd & "', '" & Request("TypeDec") & "', '" & Request("OP") & "', " & Request("Order") & ", '" & linkActive & "', " & linkObject & ", " & Align & ", '" & Dynamic & "')"
			set rs = Server.CreateObject("ADODB.RecordSet")
			set rs = conn.execute(sql)
			rI = rs(0)
			rs.close
			set rs = nothing
			
			If Request("NameTrad") <> "" Then
				SaveNewTrad Request("NameTrad"), "CartRep", "ID", "alterName", rI
			End If
				
			If Request("customSqlDef") <> "" Then
				SaveNewDef Request("customSqlDef"), rI
			End If
			
			conn.close
		Case "u"
			ArrVal = Split(Request("ID"),", ")
			For i = 0 to UBound(ArrVal)
				ID = Replace(ArrVal(i), "-", "_")
				If Request("TypeRnd" & ID) = "ON" Then TypeRnd = "Y" Else TypeRnd = "N"
				sql = sql & " update olkCartRep set Name = N'" & saveHTMLDecode(Request("Name" & ID), False) & "', Access = '" & _
				Request("Access" & ID) & "', Type = '" & Request("Type" & ID) & "', OP = '" & _
				Request("OP" & ID) & "', TypeRnd = '" & TypeRnd & "', [Order] = " & Request("Order" & ID) & " where ID = " & ArrVal(i)
			Next
			If sql <> "" Then conn.execute(sql)
		Case "del"
			sql = "delete olkCartRep where ID = " & Request("rI") & _
				  " delete olkCartRepAlterNames where ID = " & Request("rI")
			conn.execute(sql)
		Case "e"
			If Request("TypeRnd") = "ON" Then TypeRnd = "Y" Else TypeRnd = "N"
			If Request("linkActive") = "Y" Then linkActive = "Y" Else linkActive = "N"
			If Request("linkObject") <> "" Then linkObject = Request("linkObject") Else linkObject = "NULL"
			If Request("Align") <> "" Then Align = "'" & Request("Align") & "'" Else Align = "NULL"
			If Request("chkDynamic") = "Y" Then Dynamic = "Y" Else Dynamic = "N"
			sql = "update olkCartRep set Name = N'" & saveHTMLDecode(Request("Name"), False) & "', Access = '" & Request("Access") & _
				  "', Field = N'" & saveHTMLDecode(Request("customSql"), False) & "', Type = '" & Request("Type") & _
				  "', OP = '" & Request("OP") & "', TypeRnd = '" & TypeRnd & "', TypeDec = '" & Request("TypeDec") & "', " & _
				  "[Order] = " & Request("Order") & ", linkActive = '" & linkActive & "', linkObject = " & linkObject & ", Align = " & Align & ", Dynamic = '" & Dynamic & "' where ID = " & Request("rI")
			conn.execute(sql)
			rI = Request("rI")
			
			sql = "delete olkCartRepLinksVars where ID = " & rI
			conn.execute(sql)
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandText = "DBOLKAdminCartOptLinksVars" & Session("ID")
			cmd.CommandType = adCmdStoredProc
			cmd("@ID") = rI
			
			varVar = Split(Request("varVar"), ", ")
			For i = 0 to UBound(varVar)
				valBy = Request("valBy" & varVar(i))
				Select Case valBy
					Case "F"
						If Request("valValueF" & varVar(i)) <> "" Then 
							cmd("@valValue") = saveHTMLDecode(Request("valValueF" & varVar(i)), True)
						End If
					Case "V"
						If Request("varDataType" & varVar(i)) <> "datetime" Then
							If Request("valValueV" & varVar(i)) <> "" Then
								cmd("@valValue") = saveHTMLDecode(Request("valValueV" & varVar(i)), True)
							End If
						Else
							If Request("colValDat" & varVar(i)) <> "" Then
								cmd("@valDate") = SaveSqlDate(Request("colValDat" & varVar(i)))
							End If
						End If
					Case "Q"
						cmd("@valValue") = saveHTMLDecode(Request("valQuery" & varVar(i)), True)
				End Select
				cmd("@varId") = varVar(i)
				cmd("@valBy") = valBy
				cmd.execute()
			next
	End Select
	'GenMyQuery "ItemRepA"
	'GenMyQuery "ItemRepC"
	'GenMyQuery "ItemRepM"
	If Request("btnApply") <> "" Then
		response.Redirect "adminCartOpt.asp?edit=Y&rI=" & rI & "&1=1#table20"
	Else
		response.Redirect "adminCartOpt.asp"
	End If
End Sub

Private Sub adminInvOpt()
	Select Case Request("cmd")
		Case "a"
			rowField = saveHTMLDecode(Request("customSql"), False)
			If Request("rowTypeRnd") = "ON" Then rowTypeRnd = "Y" Else rowTypeRnd = "N"
			If Request("linkActive") = "Y" Then linkActive = "Y" Else linkActive = "N"
			If Request("linkObject") <> "" Then linkObject = Request("linkObject") Else linkObject = "NULL"
			If Request("chkHideNull") = "Y" Then chkHideNull = "N" Else chkHideNull = "N"
			sql = "declare @rowIndex int set @rowIndex = ISNULL((select max(rowIndex)+1 from olkItemRep),1) " & _
				  "select @rowIndex rowIndex " & _
				  "insert olkItemRep(rowIndex, rowName, rowField, rowAccess, rowType, rowTypeRnd, rowTypeDec, rowOP, rowOrder, HideNull, linkActive, linkObject) " & _
				  "values(@rowIndex,N'" & saveHTMLDecode(Request("rowName"), False) & "',N'" & rowField & "',N'" & Request("rowAccess") & _
				  "','" & Request("rowType") & "', '" & rowTypeRnd & "', '" & Request("rowTypeDec") & "', '" & Request("rowOP") & "', " & Request("RowOrder") & ", '" & chkHideNull & "', '" & linkActive & "', " & linkObject & ")"
			set rs = Server.CreateObject("ADODB.RecordSet")
			set rs = conn.execute(sql)
			rI = rs(0)
			rs.close
			set rs = nothing
			
			If Request("rowNameTrad") <> "" Then
				SaveNewTrad Request("rowNameTrad"), "ItemRep", "rowIndex", "alterRowName", rI
			End If
				
			If Request("customSqlDef") <> "" Then
				SaveNewDef Request("customSqlDef"), rI
			End If
			
			conn.close
		Case "u"
			ArrVal = Split(Request("rowIndex"),", ")
			For i = 0 to UBound(ArrVal)
				rowIndex = Replace(ArrVal(i), "-", "_")
				If Request("rowTypeRnd" & rowIndex) = "ON" Then rowTypeRnd = "Y" Else rowTypeRnd = "N"
				sql = sql & " update olkItemRep set rowName = N'" & saveHTMLDecode(Request("rowName" & rowIndex), False) & "', rowAccess = '" & _
				Request("rowAccess" & rowIndex) & "', rowType = '" & Request("rowType" & rowIndex) & "', rowOP = '" & _
				Request("rowOP" & rowIndex) & "', rowTypeRnd = '" & rowTypeRnd & "', rowOrder = " & Request("RowOrder" & rowIndex) & " where rowIndex = " & ArrVal(i)
			Next
			If sql <> "" Then conn.execute(sql)
		Case "del"
			sql = "delete olkItemRep where rowIndex = " & Request("rI") & _
				  " delete OLKItemRepAlterNames where rowIndex = " & Request("rI")
			conn.execute(sql)
		Case "e"
			If Request("rowTypeRnd") = "ON" Then rowTypeRnd = "Y" Else rowTypeRnd = "N"
			If Request("linkActive") = "Y" Then linkActive = "Y" Else linkActive = "N"
			If Request("linkObject") <> "" Then linkObject = Request("linkObject") Else linkObject = "NULL"
			If Request("chkHideNull") = "Y" Then chkHideNull = "Y" Else chkHideNull = "N"
			sql = "update olkItemRep set rowName = N'" & saveHTMLDecode(Request("rowName"), False) & "', rowAccess = '" & Request("rowAccess") & _
				  "', rowField = N'" & saveHTMLDecode(Request("customSql"), False) & "', rowType = '" & Request("rowType") & _
				  "', rowOP = '" & Request("rowOP") & "', rowTypeRnd = '" & rowTypeRnd & "', rowTypeDec = '" & Request("rowTypeDec") & "', rowOrder = " & Request("RowOrder") & ", HideNull = '" & chkHideNull & "', linkActive = '" & linkActive & "', linkObject = " & linkObject & " where rowIndex = " & Request("rI")
			conn.execute(sql)
			rI = Request("rI")
			
			sql = "delete OLKItemRepLinksVars where rowIndex = " & rI
			conn.execute(sql)
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandText = "DBOLKAdminInvOptLinksVars" & Session("ID")
			cmd.CommandType = adCmdStoredProc
			cmd("@rowIndex") = rI
			
			varVar = Split(Request("varVar"), ", ")
			For i = 0 to UBound(varVar)
				valBy = Request("valBy" & varVar(i))
				Select Case valBy
					Case "F"
						If Request("valValueF" & varVar(i)) <> "" Then 
							cmd("@valValue") = saveHTMLDecode(Request("valValueF" & varVar(i)), True)
						End If
					Case "V"
						If Request("varDataType" & varVar(i)) <> "datetime" Then
							If Request("valValueV" & varVar(i)) <> "" Then
								cmd("@valValue") = saveHTMLDecode(Request("valValueV" & varVar(i)), True)
							End If
						Else
							If Request("colValDat" & varVar(i)) <> "" Then
								cmd("@valDate") = SaveSqlDate(Request("colValDat" & varVar(i)))
							End If
						End If
					Case "Q"
						cmd("@valValue") = saveHTMLDecode(Request("valQuery" & varVar(i)), True)
				End Select
				cmd("@varId") = varVar(i)
				cmd("@valBy") = valBy
				cmd.execute()
			next
	End Select
	GenMyQuery "ItemRepA"
	GenMyQuery "ItemRepC"
	GenMyQuery "ItemRepM"
	If Request("btnApply") <> "" Then
		response.Redirect "adminInvOpt.asp?edit=Y&rI=" & rI & "&1=1#table20"
	Else
		response.Redirect "adminInvOpt.asp"
	End If
End Sub

Private Sub adminObjConfCols()
TypeID = Request("TypeID")
Select Case Request("cmd")
	Case "a"
		If Request("rowTypeRnd") = "ON" Then rowTypeRnd = "Y" Else rowTypeRnd = "N"
		If Request("linkActive") = "Y" Then linkActive = "Y" Else linkActive = "N"
		If Request("chkActive") = "Y" Then Active = "Y" Else Active = "N"
		If Request("LinkLink") <> "" Then LinkLink = "N'" & saveHTMLDecode(Request("LinkLink"), False) & "'" Else LinkLink = "NULL"
		Select Case Request("LinkType")
			Case "R"
				If Request("linkObjectRS") <> "" Then linkObject = Request("linkObjectRS") Else linkObject = "NULL"
			Case "F"
				If Request("linkObjectForm") <> "" Then linkObject = Request("linkObjectForm") Else linkObject = "NULL"
		End Select
		sql = "declare @TypeID nvarchar(100) set @TypeID = '" & TypeID & "' " & _
			  "declare @ID int set @ID = ISNULL((select max(ID)+1 from OLKObjConfCols where TypeID = @TypeID),1) " & _
			  "select @ID ID " & _
			  "insert OLKObjConfCols(TypeID, ID, Name, Query, Encode, EncodeRnd, EncodeFormat, [Order], LinkType, LinkActive, LinkObject, LinkLink, Active) " & _
			  "values(@TypeID, @ID,N'" & saveHTMLDecode(Request("rowName"), False) & "',N'" & saveHTMLDecode(Request("customSql"), False) & "', " & _
			  "'" & Request("rowType") & "', '" & rowTypeRnd & "', '" & Request("rowTypeDec") & "', " & Request("RowOrder") & ", '" & Request("LinkType") & "', '" & linkActive & "', " & linkObject & ", " & LinkLink & ", '" & Active & "')"
		set rs = Server.CreateObject("ADODB.RecordSet")
		set rs = conn.execute(sql)
		ID = rs(0)
		rs.close
		set rs = nothing
		
		If Request("rowNameTrad") <> "" Then
			SaveNewTrad Request("rowNameTrad"), "ObjConfCols", "TypeID,ID", "alterName", TypeID & "," & ID
		End If
			
		If Request("customSqlDef") <> "" Then
			SaveNewDef Request("customSqlDef"), TypeID & "," & ID
		End If
		
		GenMyQuery "ExecConf" & TypeID
				
		conn.close
	Case "u"
		ArrVal = Split(Request("rowIndex"),", ")
		For i = 0 to UBound(ArrVal)
			rowIndex = Replace(ArrVal(i), "-", "_")
			If Request("rowTypeRnd" & rowIndex) = "ON" Then rowTypeRnd = "Y" Else rowTypeRnd = "N"
			If Request("chkActive" & rowIndex) = "Y" Then Active = "Y" Else Active = "N"
			sql = sql & " update OLKObjConfCols set Name = N'" & saveHTMLDecode(Request("rowName" & rowIndex), False) & "', Encode = '" & Request("rowType" & rowIndex) & "', " & _
			"EncodeRnd = '" & rowTypeRnd & "', [Order] = " & Request("RowOrder" & rowIndex) & ", Active = '" & Active & "' where TypeID = '" & TypeID & "' and ID = " & ArrVal(i) & " "
		Next
		If sql <> "" Then conn.execute(sql)
		
		GenMyQuery "ExecConf" & TypeID
	Case "del"
		sql = "declare @TypeID nvarchar(100) set @TypeID = '" & TypeID & "' " & _
				"declare @ID int set @ID = " & Request("ID") & " " & _
				"delete OLKObjConfCols where TypeID = @TypeID and ID = @ID " & _
			  	"delete OLKObjConfColsAlterNames where TypeID = @TypeID and ID = @ID "
		conn.execute(sql)
		
		GenMyQuery "ExecConf" & TypeID
	Case "e"
		ID = Request("ID")
		If Request("rowTypeRnd") = "ON" Then rowTypeRnd = "Y" Else rowTypeRnd = "N"
		If Request("linkActive") = "Y" Then linkActive = "Y" Else linkActive = "N"
		If Request("chkActive") = "Y" Then Active = "Y" Else Active = "N"
		If Request("LinkLink") <> "" Then LinkLink = "N'" & saveHTMLDecode(Request("LinkLink"), False) & "'" Else LinkLink = "NULL"
		Select Case Request("LinkType")
			Case "R"
				If Request("linkObjectRS") <> "" Then linkObject = Request("linkObjectRS") Else linkObject = "NULL"
			Case "F"
				If Request("linkObjectForm") <> "" Then linkObject = Request("linkObjectForm") Else linkObject = "NULL"
		End Select
		sql = "update OLKObjConfCols set Name = N'" & saveHTMLDecode(Request("rowName"), False) & "', [Query] = N'" & saveHTMLDecode(Request("customSql"), False) & "', " & _
			"Encode = '" & Request("rowType") & "', EncodeRnd = '" & rowTypeRnd & "', EncodeFormat = '" & Request("rowTypeDec") & "', [Order] = " & Request("RowOrder") & ", " & _
			"LinkType = '" & Request("LinkType") & "', linkActive = '" & linkActive & "', linkObject = " & linkObject & ", LinkLink = " & LinkLink & ", Active = '" & Active & "' where TypeID = '" & TypeID & "' and ID = " & ID
		conn.execute(sql)
		
		sql = "delete OLKObjConfColsLinksVars where TypeID = '" & TypeID & "' and ID = " & ID
		conn.execute(sql)
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandText = "DBOLKAdminObjConfColsLinksVars" & Session("ID")
		cmd.CommandType = adCmdStoredProc
		cmd("@TypeID") = TypeID
		cmd("@ID") = ID
		
		varVar = Split(Request("varVar"), ", ")
		For i = 0 to UBound(varVar)
			valBy = Request("valBy" & varVar(i))
			Select Case valBy
				Case "F"
					If Request("valValueF" & varVar(i)) <> "" Then 
						cmd("@Value") = saveHTMLDecode(Request("valValueF" & varVar(i)), True)
					End If
				Case "V"
					If Request("varDataType" & varVar(i)) <> "datetime" Then
						If Request("valValueV" & varVar(i)) <> "" Then
							cmd("@Value") = saveHTMLDecode(Request("valValueV" & varVar(i)), True)
						End If
					Else
						If Request("colValDat" & varVar(i)) <> "" Then
							cmd("@Date") = SaveSqlDate(Request("colValDat" & varVar(i)))
						End If
					End If
				Case "Q"
					cmd("@Value") = saveHTMLDecode(Request("valQuery" & varVar(i)), True)
			End Select
			cmd("@VarId") = varVar(i)
			cmd("@By") = valBy
			cmd.execute()
		next
		
		GenMyQuery "ExecConf" & TypeID
End Select
If Request("btnApply") <> "" Then
	response.Redirect "adminObjConfCols.asp?TypeID=" & TypeID & "&ID=" & ID
Else
	response.Redirect "adminObjConfCols.asp?TypeID=" & TypeID
End If
End Sub

Private Sub adminCrdOpt()
Select Case Request("cmd")
	Case "a"
		rowField = saveHTMLDecode(Request("customSql"), False)
		If Request("rowTypeRnd") = "ON" Then rowTypeRnd = "Y" Else rowTypeRnd = "N"
		sql = "declare @rowIndex int set @rowIndex = ISNULL((select max(rowIndex)+1 from olkCardRep),1) " & _
			  "select @rowIndex rowIndex " & _
			  "insert olkCardRep(rowIndex, rowName, rowField, rowAccess, rowType, rowTypeRnd, rowTypeDec, rowOP, rowOrder, colIndex, ShowAt, RowAlign) " & _
			  "values(@rowIndex,N'" & saveHTMLDecode(Request("rowName"), False) & "',N'" & rowField & "',N'" & Request("rowAccess") & _
			  "','" & Request("rowType") & "', '" & rowTypeRnd & "', '" & Request("rowTypeDec") & "', '" & Request("rowOP") & "', " & Request("RowOrder") & ", '" & Request("colIndex") & "', '" & Request("showAt") & "', '" & Request("RowAlign") & "')"
		set rs = Server.CreateObject("ADODB.RecordSet")
		set rs = conn.execute(sql)
		rI = rs(0)
		rs.close
		set rs = nothing
		
		If Request("rowNameTrad") <> "" Then
			SaveNewTrad Request("rowNameTrad"), "CardRep", "rowIndex", "alterRowName", rI
		End If
			
		If Request("customSqlDef") <> "" Then
			SaveNewDef Request("customSqlDef"), rI
		End If

	Case "u"
		ArrVal = Split(Request("rowIndex"),", ")
		For i = 0 to UBound(ArrVal)
		If Request("rowTypeRnd" & ArrVal(i)) = "ON" Then rowTypeRnd = "Y" Else rowTypeRnd = "N"
		sql = sql & " update olkCardRep set rowName = N'" & saveHTMLDecode(Request("rowName" & ArrVal(i)), False) & "', rowAccess = '" & _
		Request("rowAccess" & ArrVal(i)) & "', rowType = '" & Request("rowType" & ArrVal(i)) & "', rowOP = '" & _
		Request("rowOP" & ArrVal(i)) & "', rowTypeRnd = '" & rowTypeRnd & "', rowOrder = " & Request("RowOrder" & ArrVal(i)) & ", colIndex = '" & Request("colIndex" & ArrVal(i)) & "', showAt = '" & Request("showAt" & ArrVal(i)) & "' where rowIndex = " & ArrVal(i)
		Next
		If sql <> "" Then conn.execute(sql)
	Case "del"
		sql = "delete olkCardRep where rowIndex = " & Request("rI") & _
			  " delete olkCardRepAlterNames where rowIndex = " & Request("rI")
		conn.execute(sql)
	Case "e"
		If Request("rowTypeRnd") = "ON" Then rowTypeRnd = "Y" Else rowTypeRnd = "N"
		sql = "update olkCardRep set rowName = N'" & saveHTMLDecode(Request("rowName"), False) & "', rowAccess = '" & Request("rowAccess") & _
			  "', rowField = N'" & saveHTMLDecode(Request("customSql"), False) & "', rowType = '" & Request("rowType") & _
			  "', rowOP = '" & Request("rowOP") & "', rowTypeRnd = '" & rowTypeRnd & "', rowTypeDec = '" & Request("rowTypeDec") & "', rowOrder = " & Request("RowOrder") & ", " & _
			  "colIndex = '" & Request("colIndex") & "', showAt = '" & Request("showAt") & "', RowAlign = '" & Request("RowAlign") & "' where rowIndex = " & Request("rI")
		conn.execute(sql)
		rI = Request("rI")
End Select
GenMyQuery "OCRD"
If Request("btnApply") <> "" Then
	response.Redirect "adminCardOpt.asp?edit=Y&rI=" & rI & "&1=1#table20"
Else
	response.Redirect "adminCardOpt.asp"
End If
End Sub


Private Sub adminCatOpt()

	Select Case Request("cmd")
		Case "a"
			If Request("ColTypeRnd") = "Y" Then ColTypeRnd = "Y" Else ColTypeRnd = "N"
			If Request("ReqLogin") = "Y" Then ReqLogin = "Y" Else ReqLogin = "N"
			If Request("OLKCType") = "T" Then
				sqlAdd1 = " , ColIndex"
				sqlAdd2 = ", '" & Request("ColIndex") & "'"
			End If
			If Request("ColName") <> "" Then ColName = saveHTMLDecode(Request("ColName"), False) Else ColName = Request("colField")
			ColQuery = saveHTMLDecode(Request("ColQuery"), False)
			sql = "declare @LineIndex int set @LineIndex = ISNULL((select max(LineIndex)+1 from OLK" & Request("OLKCType") & "Cart),1) select @LineIndex LineIndex " & _
				  "insert OLK" & Request("OLKCType") & "Cart(LineIndex, colName, colQuery, colAccess, colType, colTypeRnd, colTypeDec, colAlign, ColOrdr, ReqLogin" & sqlAdd1 & ") " & _
				  "values(@LineIndex,N'" & colName & "',N'" & colQuery & "',N'" & Request("colAccess") & _
				  "',N'" & Request("colType") & "', '" & ColTypeRnd & "', N'" & Request("colTypeDec") & "', '" & Request("colAlign") & "', " & Request("ColOrdr") & ", '" & ReqLogin & "' " & sqlAdd2 & ")"
			set rs = Server.CreateObject("ADODB.RecordSet")
			set rs = conn.execute(sql)
			LineIndex = rs(0)
			If Request("btnApply") <> "" Then
				RedirVal = "adminCatOpt.asp?edit=Y&LineIndex=" & LineIndex & "&OLKCType=" & Request("OLKCType")
			End If
			
			If Request("colNameTrad") <> "" Then
				SaveNewTrad Request("colNameTrad"), Request("OLKCType") & "Cart", "LineIndex", "alterColName", LineIndex
			End If
			
			Select Case Request("OLKCType")
				Case "T"
					DefID = 1
				Case "C"
					DefID = 2
			End Select
			If Request("ColQueryDef") <> "" Then
				SaveNewDef Request("ColQueryDef"), CStr(DefID) & CStr(LineIndex)
			End If

		Case "u"
			If Request("OLKCType") = "C" Then
				sqlAddC = ", catCols = " & Request("catColsC") & ", pdfCols = " & Request("pdfColsC")
				sqlAddV = ", catCols = " & Request("catColsV") & ", pdfCols = " & Request("pdfColsV")
			End If
			sql = "update OLKCatOpt set ImgMaxSize = " & Request("ImgMaxSizeC") & sqlAddC & _
			", catRows = " & Request("catRowsC") & " where UserType = 'C' and CatType = '" & Request("OLKCType") & _
			"' update OLKCatOpt set ImgMaxSize = " & Request("ImgMaxSizeV") & sqlAddV & _
			", catRows = " & Request("catRowsV") & " where UserType = 'V' and CatType = '" & Request("OLKCType") & "'"
			conn.execute(sql)

			set rs = server.createobject("ADODB.Recordset")
			sql = "select LineIndex from OLK" & Request("OLKCType") & "Cart"
			set rs = conn.execute(sql)
			sql = ""
			sqlAddStr = ""
			do while not rs.eof
				reqVar = rs("LineIndex")				
				If Request("ColTypeRnd" & reqVar) = "Y" Then ColTypeRnd = "Y" Else ColTypeRnd = "N"
				If Request("OLKCType") = "T" Then sqlAddStr = ", ColIndex = '" & Request("ColIndex" & reqVar) & "' "
				If Request("ReqLogin" & reqVar) = "Y" Then ReqLogin = "Y" Else ReqLogin = "N"
				sql = sql & "update OLK" & Request("OLKCType") & "Cart set ColName = N'" & saveHTMLDecode(Request("ColName" & reqVar), False) & "', ColType = '" & Request("ColType" & reqVar) & "', " & _
				"ColTypeRnd = '" & ColTypeRnd & "', ColAccess = '" & Request("ColAccess" & reqVar) & "', ColAlign = '" & Request("ColAlign" & reqVar) & "', ColOrdr = " & Request("ColOrdr" & reqVar) & ", ReqLogin = '" & ReqLogin & "' " & sqlAddStr & " where LineIndex = " & rs("LineIndex")
			rs.movenext
			loop
			If sql <> "" Then conn.execute(sql)
			set rs = nothing
		Case "del"
			sql = "delete OLK" & Request("OLKCType") & "Cart where LineIndex = " & Request("LineIndex") & sqlAdd & _
				" delete OLK" & Request("OLKCType") & "CartAlterNames where LineIndex = " & Request("LineIndex") & sqlAdd
			conn.execute(sql)
		Case "e"
			If Request("ColTypeRnd") = "Y" Then ColTypeRnd = "Y" Else ColTypeRnd = "N"
			If Request("ReqLogin") = "Y" Then ReqLogin = "Y" Else ReqLogin = "N"
			If Request("OLKCType") = "T" Then sqlAdd = " , ColIndex = '" & Request("ColIndex") & "' "
			sql = "update OLK" & Request("OLKCType") & "Cart set ColName = N'" & saveHTMLDecode(Request("ColName"), False) & "', ColQuery = N'" & saveHTMLDecode(Request("ColQuery"), False) & "', " & _
			"ColType = '" & Request("ColType") & "', ColTypeRnd = '" & ColTypeRnd & "', colTypeDec = '" & Request("colTypeDec") & "', ColAccess = '" & Request("ColAccess") & "', " & _
			"ColAlign = '" & Request("ColAlign") & "', ColOrdr = " & Request("ColOrdr") & ", ReqLogin = '" & ReqLogin & "' " & sqlAdd & " where LineIndex = " & Request("LineIndex")
			conn.execute(sql)
			If Request("btnApply") <> "" Then
				RedirVal = "adminCatOpt.asp?edit=Y&LineIndex=" & Request("LineIndex") & "&OLKCType=" & Request("OLKCType") 
			End If
	End Select
	conn.close
	If RedirVal = "" Then
		response.redirect "adminCatOpt.asp?OLKCType=" & Request("OLKCType")
	Else
		Response.Redirect RedirVal
	End If
End Sub

Private Sub setActiveMail
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKSetActiveMail" & Session("ID")
	cmd.execute()
End Sub

Public Sub WinClose(rVal)
doWinClose = True %>
<script language="javascript" src="general.js"></script>
<script language="javascript">
<% If rVal <> "" Then %>opener.location.href='<%=rVal%>'<% End If %>
<% If PwdChanged Then %>alert("<%=getadminSubmitLngStr("LtxtChangePwdConf")%>")<% End If %>
window.close()
</script>
<% End Sub %>

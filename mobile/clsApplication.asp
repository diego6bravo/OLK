<!--#include virtual="adovbs.inc"-->
<!--#include file="conn.asp"-->
<%
iisVer = Request.ServerVariables("SERVER_SOFTWARE")
iisVer = Right(iisVer, Len(iisVer)-InStr(iisVer, "/"))
If iisVer >= "6.0" Then Response.CodePage = 65001
Response.CharSet = "UTF-8"

If Session("olkdb") = "" Then db = "OLKCommon" Else db = Session("olkdb")

connStr = "Provider=SQLOLEDB;charset=utf8;Data Source=" & olkip & ";Initial Catalog=" & db & ";Uid=" & olklogin & ";Pwd=" & olkpass & ""
set conn = Server.CreateObject("ADODB.Connection")
set connCommon = Server.CreateObject("ADODB.Connection")
set cmd = Server.CreateObject("ADODB.Command")
set rs = Server.CreateObject("ADODB.RecordSet")
conn.open connStr
connCommon.open "Provider=SQLOLEDB;charset=utf8;Data Source=" & olkip & ";Initial Catalog=OLKCommon;Uid=" & olklogin & ";Pwd=" & olkpass & ""

cmd.ActiveConnection = connCommon
cmd.Commandtype = adCmdStoredProc

Class clsApplication
	Sub ConnectCommon
		If conn.State = 1 Then conn.close
		conn.open "Provider=SQLOLEDB;charset=utf8;Data Source=" & olkip & ";Initial Catalog=OLKCommon;Uid=" & olklogin & ";Pwd=" & olkpass & ""
	End Sub
	
	Sub ConnectDB
		ConnectDatabase Session("olkdb")
	End Sub
	
	Sub ConnectDatabase(ByVal dbName)
		If conn.State = 1 Then conn.close
		conn.open "Provider=SQLOLEDB;charset=utf8;Data Source=" & olkip & ";Initial Catalog=" & dbName & ";Uid=" & olklogin & ";Pwd=" & olkpass & ""
	End Sub
	
	Sub StartApplication
		cmd.CommandText = "OLKGetGeneralSettings"
		set rs = cmd.execute()
		Application("OVersion") = rs(0)
		Application("R3Version") = rs(1)
		Application("AllowSavePwd") = rs(2) = "Y"
		Application("ShowDbName") = rs(3) = "Y"
		Application("SingleSignOn") = rs(4) = "Y"
		Application("Started") = True
		rs.close
	End Sub
	
	Sub CheckApplicationStatus
		If Not Application("Started") Then
			myApp.StartApplication
		End If
	End Sub
	
	Sub CheckLastUpdate
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKCheckLastUpdate" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LastUpdate") = myApp.LastUpdate
		set rs = cmd.execute()
		If rs("Reload") = "Y" Then ExecLoadDBConfigData
	End Sub
	
	Function IsDBUpdated(ByVal ID)
		cmd.ActiveConnection = connCommon
		cmd.CommandText = "OLKIsDBUpd"
		cmd.Parameters.Refresh()
		cmd("@ID") = ID
		cmd.execute
		IsDBUpdated = cmd.Parameters.Item(0).value = 1
	End Function
	
	Sub ExecLoadDBConfigData
		LoadDBGeneralSettings
		
		LoadAdminGeneral
		
		LoadDecSettings
		
		LoadAdminInv
		
		LoadAdminCatProp

		LoadAdminCart
		
		LoadAdminAnonLogin
		
		LoadAdminLogos
		
		LoadAdminDocConf
		
		LoadAdminObjConf
		
		LoadClientSettings
		
		LoadAutoGen
		
		LoadVerfyOrders
	End Sub
	
	Sub LoadDBConfigData(ByVal ID)
		cmd.ActiveConnection = connCommon
		cmd.CommandText = "OLKGetDBName"
		cmd.Parameters.Refresh()
		cmd("@ID") = ID
		set rs = cmd.execute
		dbName = rs("dbName")
		Session("olkdb") = dbName
		Session("ID") = ID
		ConnectDatabase dbName
		cmd.ActiveConnection = connCommon
		
		If Not Application("ConfigLoaded_" & ID) Then 
		
			ExecLoadDBConfigData
		
			Application("ConfigLoaded_" & ID) = True
		End If
	End Sub
	
	Sub ResetLastUpdate
		SetDBConfigValue "LastUpdate", Now
	End Sub
	
	Sub LoadDecSettings
		cmd.CommandText = "DBOLKGetDecSettings" & Session("ID")
		cmd.Parameters.Refresh()
		set rs = cmd.execute
		
		SetDBConfigValue "QtyDec", rs("QtyDec")
		SetDBConfigValue "PriceDec", rs("PriceDec")
		SetDBConfigValue "PercentDec", rs("PercentDec")
		SetDBConfigValue "MeasureDec", rs("MeasureDec")
		SetDBConfigValue "SumDec", rs("SumDec")
		SetDBConfigValue "RateDec", rs("RateDec")
		rs.close
		
	End Sub
	
	Sub LoadDBGeneralSettings
		cmd.CommandText = "DBOLKGetDBSettings" & Session("ID")
		cmd.Parameters.Refresh()
		set rs = cmd.execute
		
		SetDBConfigValue "VSystem", rs("SystemVersion")
		SetDBConfigValue "LawsSet", rs("LawsSet")
		SetDBConfigValue "MainCurncy", rs("MainCurncy")
		SetDBConfigValue "MDStyle", rs("MDStyle")
		SetDBConfigValue "VatPrcnt", rs("VatPrcnt")
		SetDBConfigValue "CreditLimt", rs("CreditLimt") = "Y"
		SetDBConfigValue "SalesLimit", rs("SalesLimit") = "Y"
		SetDBConfigValue "OrderLimit", rs("OrderLimit") = "Y"
		SetDBConfigValue "DlnLimit", rs("DlnLimit") = "Y"
		SetDBConfigValue "TreePricOn", rs("TreePricOn") = "Y"
		SetDBConfigValue "DirectRate", rs("DirectRate") = "Y"
		SetDBConfigValue "SegAct", rs("SegAct") = "Y"
		SetDBConfigValue "Enable3dx", rs("Enable3dx") = 1
		SetDBConfigValue "DateFormat", rs("DateFormat")
		SetDBConfigValue "SVer", rs("SVer")
		SetDBConfigValue "SVer2007", rs("SVer2007") = "Y"
		SetDBConfigValue "SVer2005", rs("SVer2005") = "Y"
		SetDBConfigValue "ImgSavePath", rs("ImgSavePath")
		SetDBConfigValue "LastUpdate", rs("LastUpdate")
		rs.close

	End Sub
	
	Sub LoadAdminGeneral
		GetAdminQuery rs, 1, null, null
		SetDBConfigValue "SlpCode", rs("SlpCode")
		SetDBConfigValue "AddDocSlpCode", rs("AddDocSlpCode")
		SetDBConfigValue "AsignedSlp", rs("AsignedSlp") = "Y"
		SetDBConfigValue "EnableBranchs", rs("EnableBranchs") = "Y"
		SetDBConfigValue "showCxcIncTrans", rs("showCxcIncTrans") = "Y"
		SetDBConfigValue "showCxcIncTransC", rs("showCxcIncTransC") = "Y"
		SetDBConfigValue "ecdays", rs("ecdays")
		SetDBConfigValue "showCxcOpenInv", rs("showCxcOpenInv") = "Y"
		SetDBConfigValue "showCxcOpenInvC", rs("showCxcOpenInvC") = "Y"
		SetDBConfigValue "showCxcOpenInvBy", rs("showCxcOpenInvBy")
		SetDBConfigValue "showCxcOpenInvByC", rs("showCxcOpenInvByC")
		SetDBConfigValue "showCxcDueDate", rs("showCxcDueDate") = "Y"
		SetDBConfigValue "showCxcDueDateC", rs("showCxcDueDateC") = "Y"
		SetDBConfigValue "NatLng", rs("NatLng")
		SetDBConfigValue "CopyLastFCRate", rs("CopyLastFCRate") = "Y"
		SetDBConfigValue "EnRetroPoll", rs("EnRetroPoll") = "Y"
		SetDBConfigValue "AllowAgentAccessCDoc", rs("AllowAgentAccessCDoc") = "Y"
		SetDBConfigValue "DefClientOPTab", rs("DefClientOPTab")
		SetDBConfigValue "ActOrdr1", rs("ActOrdr1")
		SetDBConfigValue "ActOrdr2", rs("ActOrdr2")
		SetDBConfigValue "EnRetroPollDays", rs("EnRetroPollDays")
		SetDBConfigValue "EnBlockRClk", rs("EnBlockRClk") = "Y"
		SetDBConfigValue "EnBlkRClkMsg", rs("EnBlkRClkMsg")
		SetDBConfigValue "LawsSet", rs("LawsSet")
		SetDBConfigValue "LCID", rs("LCID")
		SetDBConfigValue "EnableCLogLogin", rs("EnableCLogLogin") = "Y"
		SetDBConfigValue "EnableVLogLogin", rs("EnableVLogLogin") = "Y"
		SetDBConfigValue "EnableCSearchFilterLog", rs("EnableCSearchFilterLog") = "Y"
		SetDBConfigValue "EnableCSearchItemLog", rs("EnableCSearchItemLog") = "Y"
		SetDBConfigValue "EnableCItemViewLog", rs("EnableCItemViewLog") = "Y"
		SetDBConfigValue "EnableCItemPurLog", rs("EnableCItemPurLog") = "Y"
		SetDBConfigValue "AgentClientsFilter", rs("AgentClientsFilter")
		SetDBConfigValue "EnableCSearchByVatId", rs("EnableCSearchByVatId") = "Y"
		SetDBConfigValue "EnableCSearchByLicTradNum", rs("EnableCSearchByLicTradNum") = "Y"
		SetDBConfigValue "AlterLocation", rs("AlterLocation") = "Y"
		SetDBConfigValue "ViewDocFilter", rs("ViewDocFilter")
		SetDBConfigValue "AddUserApr", rs("AddUserApr") = "Y"
		SetDBConfigValue "EnableAtt", rs("EnableAtt") = "Y"
		SetDBConfigValue "ShowBPCountry", rs("ShowBPCountry") = "Y"
		SetDBConfigValue "FlowAutDirectAdd", rs("FlowAutDirectAdd") = "Y"
		SetDBConfigValue "EnableDelegate", rs("EnableDelegate") = "Y"
		rs.close
	End Sub
	
	Sub LoadAdminInv
		GetAdminQuery rs, 2, null, null

		SetDBConfigValue "EnableMinInv", rs("EnableMinInv")	= "Y"
		SetDBConfigValue "MinInv", rs("MinInv")	
		SetDBConfigValue "MinInvV", rs("MinInvV")	
		SetDBConfigValue "MinInvBy", rs("MinInvBy")	
		SetDBConfigValue "EnableMinInvV", rs("EnableMinInvV") = "Y"
		SetDBConfigValue "MinInvVBy", rs("MinInvVBy")	
		SetDBConfigValue "MinPrice", rs("MinPrice")	
		SetDBConfigValue "VerfyDisp", rs("VerfyDisp")	
		SetDBConfigValue "VerfyDispWhs", rs("VerfyDispWhs")	
		SetDBConfigValue "VerfyDispMethod", rs("VerfyDispMethod")	
		SetDBConfigValue "WhsCode", rs("WhsCode")	
		SetDBConfigValue "InvBDGBy", rs("InvBDGBy")	
		SetDBConfigValue "f_creacion", rs("f_creacion")	
		SetDBConfigValue "ManageItmWhs", rs("ManageItmWhs")	= "Y"
		SetDBConfigValue "VerfyGo2", rs("VerfyGo2")	
		SetDBConfigValue "GenFilter", rs("GenFilter")	
		SetDBConfigValue "GenFAppV", rs("GenFAppV")	= "Y"
		SetDBConfigValue "GenFAppC", rs("GenFAppC")	= "Y"
		SetDBConfigValue "EnableCodeBarsQry", rs("EnableCodeBarsQry") = "Y"
		SetDBConfigValue "CodeBarsQryMethod", rs("CodeBarsQryMethod")
		SetDBConfigValue "CodeBarsQry", rs("CodeBarsQry")	
		SetDBConfigValue "ApplyInvFiltersBy", rs("ApplyInvFiltersBy")	
		SetDBConfigValue "EnableItemRec", rs("EnableItemRec") = "Y"
		SetDBConfigValue "ItemRecQry", rs("ItemRecQry")	
		SetDBConfigValue "EnableCombos", rs("EnableCombos") = "Y"
		SetDBConfigValue "EnableSearchItmSupp", rs("EnableSearchItmSupp") = "Y"
		rs.close
	End Sub
	
	Sub LoadAdminCatProp
		GetAdminQuery rs, 4, null, null
		SetDBConfigValue "DefCatOrdrC", rs("DefCatOrdrC")	
		SetDBConfigValue "DefCatOrdrV", rs("DefCatOrdrV")	
		SetDBConfigValue "AutoSearchOpen", rs("AutoSearchOpen")	
		SetDBConfigValue "ShowClientRef", rs("ShowClientRef") = "Y"
		SetDBConfigValue "CatShowProm", rs("CatShowProm")	
		SetDBConfigValue "ShowClientSalUn", rs("ShowClientSalUn") = "Y"
		SetDBConfigValue "ShowPocketImg", rs("ShowPocketImg") = "Y"
		SetDBConfigValue "ShowClientImg", rs("ShowClientImg") = "Y"
		SetDBConfigValue "ShowAgentImg", rs("ShowAgentImg") = "Y"
		SetDBConfigValue "AgentSaleUnit", rs("AgentSaleUnit")	
		SetDBConfigValue "CarArt", rs("CarArt")	
		SetDBConfigValue "ClientSaleUnit", rs("ClientSaleUnit")	
		SetDBConfigValue "UnEmbPriceSet", rs("UnEmbPriceSet") = "Y"
		SetDBConfigValue "olkItemReport2", rs("olkItemReport2")	
		SetDBConfigValue "EnableOfertToDisc", rs("EnableOfertToDisc") = "Y"
		SetDBConfigValue "ShowNotAvlInv", rs("ShowNotAvlInv") = "Y"
		SetDBConfigValue "ShowSearchTreeCount", rs("ShowSearchTreeCount") = "Y"
		SetDBConfigValue "ShowSearchTreeSubCount", rs("ShowSearchTreeSubCount") = "Y"
		SetDBConfigValue "EnableSearchAlterCode", rs("EnableSearchAlterCode") = "Y"
		SetDBConfigValue "DefViewCL", rs("DefViewCL")	
		SetDBConfigValue "DefViewAG", rs("DefViewAG")
		SetDBConfigValue "ShowQtyInUnAg", rs("ShowQtyInUnAg") = "Y"
		SetDBConfigValue "ShowQtyInUnCl", rs("ShowQtyInUnCl") = "Y"
		SetDBConfigValue "SearchExactA", rs("SearchExactA") = "Y"
		SetDBConfigValue "SearchMethodA", rs("SearchMethodA")	
		SetDBConfigValue "SearchExactC", rs("SearchExactC") = "Y"
		SetDBConfigValue "SearchMethodC", rs("SearchMethodC")	
		SetDBConfigValue "SearchExactP", rs("SearchExactP") = "Y"
		SetDBConfigValue "SearchMethodP", rs("SearchMethodP")	
		SetDBConfigValue "ShowPriceTax", rs("ShowPriceTax") = "Y"
		SetDBConfigValue "SearchByVendorCode", rs("SearchByVendorCode") = "Y"
		SetDBConfigValue "EnableUnitSelection", rs("EnableUnitSelection") = "Y"
		SetDBConfigValue "EnableMultCheck", rs("EnableMultCheck") = "Y"
		rs.close
	End Sub
	
	Sub LoadAdminCart
		GetAdminQuery rs, 5, null, null
		SetDBConfigValue "CartSumQty", CInt(rs("CartSumQty"))
		SetDBConfigValue "EnableCartSum", rs("EnableCartSum") = "Y"
		SetDBConfigValue "Top10Items", rs("Top10Items")
		SetDBConfigValue "EnTop10Items", rs("EnTop10Items") = "Y"
		SetDBConfigValue "ExpItems", rs("ExpItems") = "Y"
		SetDBConfigValue "DocMCBal", rs("DocMCBal")
		SetDBConfigValue "CartGroup", rs("CartGroup")
		SetDBConfigValue "SDKLineMemo", rs("SDKLineMemo") = "Y"
		SetDBConfigValue "D_DocC", rs("D_DocC")
		SetDBConfigValue "PocketDefDoc", rs("PocketDefDoc")
		SetDBConfigValue "EnableClientMDoc", rs("EnableClientMDoc") = "Y"
		SetDBConfigValue "AfterCartAddC", rs("AfterCartAddC")
		SetDBConfigValue "AfterCartAddV", rs("AfterCartAddV")
		SetDBConfigValue "AfterCartAddPocket", rs("AfterCartAddPocket")
		SetDBConfigValue "BasketMItems", rs("BasketMItems") = "Y"
		SetDBConfigValue "EnCSelDoc", rs("EnCSelDoc") = "Y"
		SetDBConfigValue "CCartNote", rs("CCartNote")
		SetDBConfigValue "PrintCCartNote", rs("PrintCCartNote") = "Y"
		SetDBConfigValue "EnableCartImpC", rs("EnableCartImpC") = "Y"
		SetDBConfigValue "EnableCartImpV", rs("EnableCartImpV") = "Y"
		SetDBConfigValue "EnableDiscount", rs("EnableDiscount") = "Y"
		SetDBConfigValue "ShowPriceBefDiscount", rs("ShowPriceBefDiscount") = "Y"
		SetDBConfigValue "MaxDiscount", rs("MaxDiscount")
		SetDBConfigValue "ApplyMaxDiscToSU", rs("ApplyMaxDiscToSU") = "Y"
		SetDBConfigValue "CartType", rs("CartType")
		SetDBConfigValue "UseCustomTransMsg", rs("UseCustomTransMsg") = "Y"
		SetDBConfigValue "CustomTransMsg", rs("CustomTransMsg")
		SetDBConfigValue "ShowLineDiscount", rs("ShowLineDiscount") = "Y"
		SetDBConfigValue "PrintPriceBefDiscount", rs("PrintPriceBefDiscount") = "Y"
		SetDBConfigValue "PrintLineDiscount", rs("PrintLineDiscount") = "Y"
		SetDBConfigValue "AllowClientPartSuppSel", rs("AllowClientPartSuppSel") = "Y"
		SetDBConfigValue "EnSelAll", rs("EnSelAll") = "Y"
		SetDBConfigValue "EnSellAllUnitFrom", rs("EnSellAllUnitFrom")
		SetDBConfigValue "EnableHideCartHdr", rs("EnableHideCartHdr") = "Y"
		SetDBConfigValue "EnableDocPrjSel", rs("EnableDocPrjSel") = "Y"
		SetDBConfigValue "EnableAnonCart", rs("EnableAnonCart") = "Y"
		SetDBConfigValue "AnonCartClient", rs("AnonCartClient")
		SetDBConfigValue "FastAddUnRem", rs("FastAddUnRem") = "Y"
		SetDBConfigValue "FastAddBeep", rs("FastAddBeep") = "Y"
		SetDBConfigValue "ItemDescModQry", rs("ItemDescModQry")
		SetDBConfigValue "EnableMultiBPCart", rs("EnableMultiBPCart") = "Y"
		SetDBConfigValue "CartItmBarCode", rs("CartItmBarCode") = "Y"
		rs.close
	End Sub

	Sub LoadAdminAnonLogin
		GetAdminQuery rs, 6, null, null
		SetDBConfigValue "EnableAnSesion", rs("EnableAnSesion") = "Y"
		SetDBConfigValue "EnableAnReg", rs("EnableAnReg") = "Y"
		SetDBConfigValue "WebAddress", rs("WebAddress")	
		SetDBConfigValue "AnSesListNum", rs("AnSesListNum")	
		SetDBConfigValue "AnonSesFilter", rs("AnonSesFilter")	
		SetDBConfigValue "EnableAnRegTerms", rs("EnableAnRegTerms")	= "Y"
		SetDBConfigValue "AnTerms", rs("AnTerms")	
		SetDBConfigValue "AnRegConfAsignSLP", rs("AnRegConfAsignSLP") = "Y"
		SetDBConfigValue "EnChooseCType", rs("EnChooseCType") = "Y"
		SetDBConfigValue "AnRegConfFrom", rs("AnRegConfFrom")	
		SetDBConfigValue "AnRegConfTo", rs("AnRegConfTo")	
		SetDBConfigValue "ClientType", rs("ClientType")	
		SetDBConfigValue "AnRegAct", rs("AnRegAct")	
		SetDBConfigValue "RegActMailAdd", rs("RegActMailAdd")	
		SetDBConfigValue "RemPwdMailAdd", rs("RemPwdMailAdd")	
		SetDBConfigValue "AnRegConfRejNote", rs("AnRegConfRejNote") = "Y"
		rs.close
	End Sub
	
	Sub LoadAdminLogos	
		GetAdminQuery rs, 7, null, null
		SetDBConfigValue "TopLogo", rs("TopLogo")	
		SetDBConfigValue "MailLogo", rs("MailLogo")	
		SetDBConfigValue "AgentLogo", rs("AgentLogo")
		rs.close
	End Sub
	
	Sub LoadAdminDocConf
		GetAdminQuery rs, 9, null, null
		SetDBConfigValue "ClientReservedInvoice", rs("ClientReservedInvoice") = "Y"
		SetDBConfigValue "EnResInv", rs("EnResInv") = "Y"
		SetDBConfigValue "DefResInv", rs("DefResInv") = "Y"
		SetDBConfigValue "ORCTContraComp", rs("ORCTContraComp") = "Y"
		SetDBConfigValue "ApplyOpenRctToInvBal", rs("ApplyOpenRctToInvBal") = "Y"
		SetDBConfigValue "ChecksFilter", rs("ChecksFilter")
		SetDBConfigValue "IgnoreSystemChecksFilter", rs("IgnoreSystemChecksFilter") = "Y"
		rs.close
	End Sub
	
	Sub LoadAdminObjConf
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKGetData" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@ID") = 1
		set rs = cmd.execute()
		do while not rs.eof
			ObjID = rs("ObjectCode")
			SetDBConfigValue "ActiveObjectA" & ObjID, rs("Active") = "Y"
			SetDBConfigValue "ActiveObjectC" & ObjID, rs("ActiveClient") = "Y"
		rs.movenext
		loop
		rs.close
	End Sub
	
	Sub LoadClientSettings
		GetAdminQuery rs, 10, null, null
		SetDBConfigValue "EnableDROnlyNote", rs("EnableDROnlyNote")	= "Y"
		SetDBConfigValue "MyDataReadOnly", rs("MyDataReadOnly")	= "Y"
		rs.close
	End Sub
	
	Sub LoadAutoGen
		GetAdminQuery rs, 11, null, null
		SetDBConfigValue "AutoGenOCRD", rs("AutoGenOCRD")	= "Y"
		SetDBConfigValue "AutoGenOITM", rs("AutoGenOITM")	= "Y"
		rs.close
	End Sub
	
	Sub LoadVerfyOrders
		GetAdminQuery rs, 8, null, null
		SetDBConfigValue "VerfyBtchOrder", rs("VerfyBtch")	
		SetDBConfigValue "Verfy3dxOrder", rs("Verfy3dx")
		rs.close
	End Sub
	
	'General Functions
	
	'Returns CSS file name final string
	Function GetCSSType
		If InStr(Request.ServerVariables("HTTP_USER_AGENT"), "MSIE") <> 0 Then
			GetCSSType = "ie"
		Else
			GetCSSType = "nc"
		End If
	End Function
	
	'Returns if CSS must be Right to Left or Left to Right
	Function GetStrRtlLtr
		If Session("LanID") = 3 Then GetStrRtlLtr = "Rtl" Else GetStrRtlLtr = "Ltr"
	End Function 
	
	'Return Y/N String
	Function GetYesNo(ByVal Value)
		If Value = "Y" Then GetYesNo = "Y" Else GetYesNo = "N"
	End Function
	
	'Encodes invalid characters for forms
	Function myHTMLEncode(strVal)
		If Not IsNull(strVal) Then
			strVal = Replace(strVal, "&lt;", "<")
			strVal = Replace(strVal, "&gt;", ">")
			strVal = Replace(strVal, "&amp;", "&")
			strVal = Replace(strVal, "&quot;", """")
			myHTMLEncode = strVal
		Else
			myHTMLEncode = ""
		End If
	End Function
	
	'String Format
	Function StringFormat(sVal, aArgs)
		Dim i
		For i=0 To UBound(aArgs)
			sVal = Replace(sVal,"{" & CStr(i) & "}",aArgs(i))
		Next
	StringFormat = sVal
	End Function
	
	'Sets Database Configuration Value
	Public Sub SetDBConfigValue(ByVal ConfigID, ByVal Value)
		Application(ConfigID & "_" & Session("ID")) = Value
	End Sub
	
	'Gets Database Configuration Value
	Private Function GetDBConfigValue(ByVal ConfigID)
		GetDBConfigValue = Application(ConfigID & "_" & Session("ID"))
	End Function
	
	'Application Properties
	
	Public Property Get OLKVersion
		OLKVersion = Application("OVersion")
	End Property
	
	Public Property Get R3Version
		R3Version = Application("R3Version")
	End Property

	'General Settings
	Public Property Get AllowSavePwd
		AllowSavePwd = Application("AllowSavePwd")
	End Property
	
	Public Property Get ShowDbName
		ShowDbName = Application("ShowDbName")
	End Property
	
	Public Property Get SingleSignOn
		SingleSignOn = Application("SingleSignOn")
	End Property

	'General Settings

	Public Property Get VSystem
		VSystem = GetDBConfigValue("VSystem") 
	End Property
	
	Public Property Get QtyDec
		QtyDec = GetDBConfigValue("QtyDec") 
	End Property
	
	Public Property Get PriceDec
		PriceDec = GetDBConfigValue("PriceDec") 
	End Property
	
	Public Property Get PercentDec
		PercentDec = GetDBConfigValue("PercentDec") 
	End Property
	
	Public Property Get MeasureDec
		MeasureDec = GetDBConfigValue("MeasureDec") 
	End Property
	
	Public Property Get SumDec
		SumDec = GetDBConfigValue("SumDec") 
	End Property
	
	Public Property Get RateDec
		RateDec = GetDBConfigValue("RateDec") 
	End Property
	
	Public Property Get LawsSet
		LawsSet = GetDBConfigValue("LawsSet") 
	End Property
	
	Public Property Get MainCur
		MainCur = GetDBConfigValue("MainCurncy") 
	End Property
	
	Public Property Get MDStyle
		MDStyle = GetDBConfigValue("MDStyle") 
	End Property
	
	Public Property Get VatPrcnt
		VatPrcnt = GetDBConfigValue("VatPrcnt") 
	End Property
	
	Public Property Get CreditLimt
		CreditLimt = GetDBConfigValue("CreditLimt") 
	End Property
	
	Public Property Get SalesLimit
		SalesLimit = GetDBConfigValue("SalesLimit") 
	End Property
	
	Public Property Get OrderLimit
		OrderLimit = GetDBConfigValue("OrderLimit") 
	End Property
	
	Public Property Get DlnLimit
		DlnLimit = GetDBConfigValue("DlnLimit") 
	End Property
	
	Public Property Get TreePricOn
		TreePricOn = GetDBConfigValue("TreePricOn") 
	End Property
	
	Public Property Get DirectRate
		DirectRate = GetDBConfigValue("DirectRate") 
	End Property
	
	Public Property Get SegAct
		SegAct = GetDBConfigValue("SegAct") 
	End Property
	
	Public Property Get Enable3dx
		Enable3dx = GetDBConfigValue("Enable3dx") 
	End Property
 
	Public Property Get DateFormat
		DateFormat = GetDBConfigValue("DateFormat") 
	End Property
	
	Public Property Get SVer
		SVer = GetDBConfigValue("SVer") 
	End Property
 
	Public Property Get SVer2007
		SVer2007 = GetDBConfigValue("SVer2007") 
	End Property
 
	Public Property Get SVer2005
		SVer2005 = GetDBConfigValue("SVer2007") 
	End Property
 
	Public Property Get ImgSavePath
		ImgSavePath = GetDBConfigValue("ImgSavePath") 
	End Property
 
	Public Property Get LastUpdate
		LastUpdate = GetDBConfigValue("LastUpdate") 
	End Property
	
	'Admin General
	
	Public Property Get AddUserApr
		AddUserApr = GetDBConfigValue("AddUserApr") 
	End Property
	
	Public Property Get EnableAtt
		EnableAtt = GetDBConfigValue("EnableAtt") 
	End Property
	
	Public Property Get FlowAutDirectAdd
		FlowAutDirectAdd = GetDBConfigValue("FlowAutDirectAdd") 
	End Property
	
	Public Property Get EnableDelegate
		EnableDelegate = GetDBConfigValue("EnableDelegate") 
	End Property
	
	Public Property Get ShowBPCountry
		ShowBPCountry = GetDBConfigValue("ShowBPCountry") 
	End Property
	
	Public Property Get SlpCode
		SlpCode = GetDBConfigValue("SlpCode") 
	End Property
	
	Public Property Get AddDocSlpCode
		AddDocSlpCode = GetDBConfigValue("AddDocSlpCode") 
	End Property
	
	Public Property Get AsignedSlp
		AsignedSlp = GetDBConfigValue("AsignedSlp") 
	End Property
	
	Public Property Get EnableBranchs
		EnableBranchs = GetDBConfigValue("EnableBranchs") 
	End Property
	
	Public Property Get showCxcIncTrans
		showCxcIncTrans = GetDBConfigValue("showCxcIncTrans") 
	End Property
	
	Public Property Get showCxcIncTransC
		showCxcIncTransC = GetDBConfigValue("showCxcIncTransC") 
	End Property
	
	Public Property Get ecdays
		ecdays = GetDBConfigValue("ecdays") 
	End Property
	
	Public Property Get showCxcOpenInv
		showCxcOpenInv = GetDBConfigValue("showCxcOpenInv") 
	End Property
	
	Public Property Get showCxcOpenInvC
		showCxcOpenInvC = GetDBConfigValue("showCxcOpenInvC") 
	End Property
	
	Public Property Get showCxcOpenInvBy
		showCxcOpenInvBy = GetDBConfigValue("showCxcOpenInvBy") 
	End Property
	
	Public Property Get showCxcOpenInvByC
		showCxcOpenInvByC = GetDBConfigValue("showCxcOpenInvByC") 
	End Property
	
	Public Property Get showCxcDueDate
		showCxcDueDate = GetDBConfigValue("showCxcDueDate") 
	End Property
	
	Public Property Get showCxcDueDateC
		showCxcDueDateC = GetDBConfigValue("showCxcDueDateC") 
	End Property
	
	Public Property Get NatLng
		NatLng = GetDBConfigValue("NatLng") 
	End Property
	
	Public Property Get CopyLastFCRate
		CopyLastFCRate = GetDBConfigValue("CopyLastFCRate") 
	End Property
	
	Public Property Get EnRetroPoll
		EnRetroPoll = GetDBConfigValue("EnRetroPoll") 
	End Property
	
	Public Property Get AllowAgentAccessCDoc
		AllowAgentAccessCDoc = GetDBConfigValue("AllowAgentAccessCDoc") 
	End Property
	
	Public Property Get DefClientOPTab
		DefClientOPTab = GetDBConfigValue("DefClientOPTab") 
	End Property
	
	Public Property Get ActOrdr1
		ActOrdr1 = GetDBConfigValue("ActOrdr1") 
	End Property
	
	Public Property Get ActOrdr2
		ActOrdr2 = GetDBConfigValue("ActOrdr2") 
	End Property
	
	Public Property Get EnRetroPollDays
		EnRetroPollDays = GetDBConfigValue("EnRetroPollDays") 
	End Property
	
	Public Property Get EnBlockRClk
		EnBlockRClk = GetDBConfigValue("EnBlockRClk") 
	End Property
	
	Public Property Get EnBlkRClkMsg
		EnBlkRClkMsg = GetDBConfigValue("EnBlkRClkMsg") 
	End Property

	Public Property Get LCID
		LCID = GetDBConfigValue("LCID") 
	End Property
	
	Public Property Get EnableCLogLogin
		EnableCLogLogin = GetDBConfigValue("EnableCLogLogin") 
	End Property
	
	Public Property Get EnableVLogLogin
		EnableVLogLogin = GetDBConfigValue("EnableVLogLogin") 
	End Property
	
	Public Property Get EnableCSearchFilterLog
		EnableCSearchFilterLog = GetDBConfigValue("EnableCSearchFilterLog") 
	End Property
	
	Public Property Get EnableCSearchItemLog
		EnableCSearchItemLog = GetDBConfigValue("EnableCSearchItemLog") 
	End Property
	
	Public Property Get EnableCItemViewLog
		EnableCItemViewLog = GetDBConfigValue("EnableCItemViewLog") 
	End Property
	
	Public Property Get EnableCItemPurLog
		EnableCItemPurLog = GetDBConfigValue("EnableCItemPurLog") 
	End Property
	
	Public Property Get AgentClientsFilter
		AgentClientsFilter = GetDBConfigValue("AgentClientsFilter") 
	End Property
	
	Public Property Get EnableCSearchByVatId
		EnableCSearchByVatId = GetDBConfigValue("EnableCSearchByVatId") 
	End Property
	
	Public Property Get EnableCSearchByLicTradNum
		EnableCSearchByLicTradNum = GetDBConfigValue("EnableCSearchByLicTradNum") 
	End Property
	
	Public Property Get AlterLocation
		AlterLocation = GetDBConfigValue("AlterLocation") 
	End Property
	
	Public Property Get ViewDocFilter
		ViewDocFilter = GetDBConfigValue("ViewDocFilter") 
	End Property

	'Admin Inventory
	
	Public Property Get EnableSearchItmSupp
		EnableSearchItmSupp = GetDBConfigValue("EnableSearchItmSupp") 
	End Property
	
	Public Property Get EnableMinInv
		EnableMinInv = GetDBConfigValue("EnableMinInv") 
	End Property
	
	Public Property Get MinInv
		MinInv = GetDBConfigValue("MinInv") 
	End Property
	
	Public Property Get MinInvV
		MinInvV = GetDBConfigValue("MinInvV") 
	End Property
	
	Public Property Get MinInvBy
		MinInvBy = GetDBConfigValue("MinInvBy") 
	End Property
	
	Public Property Get EnableMinInvV
		EnableMinInvV = GetDBConfigValue("EnableMinInvV") 
	End Property
	
	Public Property Get MinInvVBy
		MinInvVBy = GetDBConfigValue("MinInvVBy") 
	End Property
	
	Public Property Get MinPrice
		MinPrice = GetDBConfigValue("MinPrice") 
	End Property
	
	Public Property Get VerfyDisp
		VerfyDisp = GetDBConfigValue("VerfyDisp") 
	End Property
	
	Public Property Get VerfyDispWhs
		VerfyDispWhs = GetDBConfigValue("VerfyDispWhs") 
	End Property
	
	Public Property Get VerfyDispMethod
		VerfyDispMethod = GetDBConfigValue("VerfyDispMethod") 
	End Property
	
	Public Property Get WhsCode
		WhsCode = GetDBConfigValue("WhsCode") 
	End Property
	
	Public Property Get InvBDGBy
		InvBDGBy = GetDBConfigValue("InvBDGBy") 
	End Property
	
	Public Property Get f_creacion
		f_creacion = GetDBConfigValue("f_creacion") 
	End Property
	
	Public Property Get ManageItmWhs
		ManageItmWhs = GetDBConfigValue("ManageItmWhs") 
	End Property
	
	Public Property Get VerfyGo2
		VerfyGo2 = GetDBConfigValue("VerfyGo2") 
	End Property
	
	Public Property Get GenFilter
		GenFilter = GetDBConfigValue("GenFilter") 
	End Property
	
	Public Property Get GenFAppV
		GenFAppV = GetDBConfigValue("GenFAppV") 
	End Property
	
	Public Property Get GenFAppC
		GenFAppC = GetDBConfigValue("GenFAppC") 
	End Property
	
	Public Property Get EnableCodeBarsQry
		EnableCodeBarsQry = GetDBConfigValue("EnableCodeBarsQry") 
	End Property
	
	Public Property Get CodeBarsQryMethod
		CodeBarsQryMethod = GetDBConfigValue("CodeBarsQryMethod") 
	End Property
	
	Public Property Get CodeBarsQry
		CodeBarsQry = GetDBConfigValue("CodeBarsQry") 
	End Property
	
	Public Property Get ApplyInvFiltersBy
		ApplyInvFiltersBy = GetDBConfigValue("ApplyInvFiltersBy") 
	End Property
	
	Public Property Get EnableItemRec
		EnableItemRec = GetDBConfigValue("EnableItemRec") 
	End Property
	
	Public Property Get ItemRecQry
		ItemRecQry = GetDBConfigValue("ItemRecQry") 
	End Property
	
	Public Property Get EnableCombos
		EnableCombos = GetDBConfigValue("EnableCombos") 
	End Property

	'Admin Cat Propeties
	
	Public Property Get EnableMultCheck
		EnableMultCheck = GetDBConfigValue("EnableMultCheck") 
	End Property
	
	Public Property Get EnableUnitSelection
		EnableUnitSelection = GetDBConfigValue("EnableUnitSelection") 
	End Property
	
	Public Property Get SearchByVendorCode
		SearchByVendorCode = GetDBConfigValue("SearchByVendorCode") 
	End Property
	
	Public Property Get ShowQtyInUnAg
		ShowQtyInUnAg = GetDBConfigValue("ShowQtyInUnAg") 
	End Property
	
	Public Property Get ShowQtyInUnCl
		ShowQtyInUnCl = GetDBConfigValue("ShowQtyInUnCl") 
	End Property
	
	Public Property Get DefCatOrdrC
		DefCatOrdrC = GetDBConfigValue("DefCatOrdrC") 
	End Property
	
	Public Property Get DefCatOrdrV
		DefCatOrdrV = GetDBConfigValue("DefCatOrdrV") 
	End Property
	
	Public Property Get AutoSearchOpen
		AutoSearchOpen = GetDBConfigValue("AutoSearchOpen") 
	End Property
	
	Public Property Get ShowClientRef
		ShowClientRef = GetDBConfigValue("ShowClientRef") 
	End Property
	
	Public Property Get CatShowProm
		CatShowProm = GetDBConfigValue("CatShowProm") 
	End Property
	
	Public Property Get ShowClientSalUn
		ShowClientSalUn = GetDBConfigValue("ShowClientSalUn") 
	End Property
	
	Public Property Get ShowPocketImg
		ShowPocketImg = GetDBConfigValue("ShowPocketImg") 
	End Property
	
	Public Property Get ShowClientImg
		ShowClientImg = GetDBConfigValue("ShowClientImg") 
	End Property
	
	Public Property Get ShowAgentImg
		ShowAgentImg = GetDBConfigValue("ShowAgentImg") 
	End Property
	
	Public Property Get AgentSaleUnit
		AgentSaleUnit = GetDBConfigValue("AgentSaleUnit") 
	End Property
	
	Public Property Get CarArt
		CarArt = GetDBConfigValue("CarArt") 
	End Property
	
	Public Property Get ClientSaleUnit
		ClientSaleUnit = GetDBConfigValue("ClientSaleUnit") 
	End Property
	
	Public Property Get UnEmbPriceSet
		UnEmbPriceSet = GetDBConfigValue("UnEmbPriceSet") 
	End Property
	
	Public Property Get olkItemReport2
		olkItemReport2 = GetDBConfigValue("olkItemReport2") 
	End Property
	
	Public Property Get EnableOfertToDisc
		EnableOfertToDisc = GetDBConfigValue("EnableOfertToDisc") 
	End Property
	
	Public Property Get ShowNotAvlInv
		ShowNotAvlInv = GetDBConfigValue("ShowNotAvlInv") 
	End Property
	
	Public Property Get ShowSearchTreeCount
		ShowSearchTreeCount = GetDBConfigValue("ShowSearchTreeCount") 
	End Property
	
	Public Property Get ShowSearchTreeSubCount
		ShowSearchTreeSubCount = GetDBConfigValue("ShowSearchTreeSubCount") 
	End Property
	
	Public Property Get EnableSearchAlterCode
		EnableSearchAlterCode = GetDBConfigValue("EnableSearchAlterCode") 
	End Property
	
	Public Property Get DefViewCL
		DefViewCL = GetDBConfigValue("DefViewCL") 
	End Property
	
	Public Property Get DefViewAG
		DefViewAG = GetDBConfigValue("DefViewAG") 
	End Property
	
	Public Property Get SearchExactA
		SearchExactA = GetDBConfigValue("SearchExactA") 
	End Property
	
	Public Property Get SearchMethodA
		SearchMethodA = GetDBConfigValue("SearchMethodA") 
	End Property
	
	Public Property Get SearchExactC
		SearchExactC = GetDBConfigValue("SearchExactC") 
	End Property
	
	Public Property Get SearchMethodC
		SearchMethodC = GetDBConfigValue("SearchMethodC") 
	End Property
	
	Public Property Get SearchExactP
		SearchExactP = GetDBConfigValue("SearchExactP") 
	End Property
	
	Public Property Get SearchMethodP
		SearchMethodP = GetDBConfigValue("SearchMethodP") 
	End Property
	
	Public Property Get ShowPriceTax
		ShowPriceTax = GetDBConfigValue("ShowPriceTax") 
	End Property
	
	'Admin Cart

	Public Property Get EnableMultiBPCart
		EnableMultiBPCart = GetDBConfigValue("EnableMultiBPCart") 
	End Property

	Public Property Get CartItmBarCode
		CartItmBarCode = GetDBConfigValue("CartItmBarCode") 
	End Property

	Public Property Get ItemDescModQry
		ItemDescModQry = GetDBConfigValue("ItemDescModQry") 
	End Property

	Public Property Get FastAddBeep
		FastAddBeep = GetDBConfigValue("FastAddBeep") 
	End Property

	Public Property Get FastAddUnRem
		FastAddUnRem = GetDBConfigValue("FastAddUnRem") 
	End Property

	Public Property Get EnableAnonCart
		EnableAnonCart = GetDBConfigValue("EnableAnonCart") 
	End Property
	
	Public Property Get AnonCartClient
		AnonCartClient = GetDBConfigValue("AnonCartClient") 
	End Property
	
	Public Property Get EnableDocPrjSel
		EnableDocPrjSel = GetDBConfigValue("EnableDocPrjSel") 
	End Property

	Public Property Get EnableHideCartHdr
		EnableHideCartHdr = GetDBConfigValue("EnableHideCartHdr") 
	End Property

	Public Property Get CartSumQty
		CartSumQty = GetDBConfigValue("CartSumQty") 
	End Property
	
	Public Property Get EnableCartSum
		EnableCartSum = GetDBConfigValue("EnableCartSum") 
	End Property
	
	Public Property Get Top10Items
		Top10Items = GetDBConfigValue("Top10Items") 
	End Property
	
	Public Property Get EnTop10Items
		EnTop10Items = GetDBConfigValue("EnTop10Items") 
	End Property
	
	Public Property Get ExpItems
		ExpItems = GetDBConfigValue("ExpItems") 
	End Property
	
	Public Property Get DocMCBal
		DocMCBal = GetDBConfigValue("DocMCBal") 
	End Property
	
	Public Property Get CartGroup
		CartGroup = GetDBConfigValue("CartGroup") 
	End Property
	
	Public Property Get SDKLineMemo
		SDKLineMemo = GetDBConfigValue("SDKLineMemo") 
	End Property
	
	Public Property Get D_DocC
		D_DocC = GetDBConfigValue("D_DocC") 
	End Property
	
	Public Property Get PocketDefDoc
		PocketDefDoc = GetDBConfigValue("PocketDefDoc") 
	End Property
	
	Public Property Get EnableClientMDoc
		EnableClientMDoc = GetDBConfigValue("EnableClientMDoc") 
	End Property
	
	Public Property Get AfterCartAddC
		AfterCartAddC = GetDBConfigValue("AfterCartAddC") 
	End Property
	
	Public Property Get AfterCartAddV
		AfterCartAddV = GetDBConfigValue("AfterCartAddV") 
	End Property
	
	Public Property Get AfterCartAddPocket
		AfterCartAddPocket = GetDBConfigValue("AfterCartAddPocket") 
	End Property
	
	Public Property Get BasketMItems
		BasketMItems = GetDBConfigValue("BasketMItems") 
	End Property
	
	Public Property Get EnCSelDoc
		EnCSelDoc = GetDBConfigValue("EnCSelDoc") 
	End Property
	
	Public Property Get CCartNote
		CCartNote = GetDBConfigValue("CCartNote") 
	End Property
	
	Public Property Get PrintCCartNote
		PrintCCartNote = GetDBConfigValue("PrintCCartNote") 
	End Property
	
	Public Property Get EnableCartImpC
		EnableCartImpC = GetDBConfigValue("EnableCartImpC") 
	End Property
	
	Public Property Get EnableCartImpV
		EnableCartImpV = GetDBConfigValue("EnableCartImpV") 
	End Property
	
	Public Property Get EnableDiscount
		EnableDiscount = GetDBConfigValue("EnableDiscount") 
	End Property
	
	Public Property Get ShowPriceBefDiscount
		ShowPriceBefDiscount = GetDBConfigValue("ShowPriceBefDiscount") 
	End Property
	
	Public Property Get MaxDiscount
		MaxDiscount = GetDBConfigValue("MaxDiscount") 
	End Property
	
	Public Property Get ApplyMaxDiscToSU
		ApplyMaxDiscToSU = GetDBConfigValue("ApplyMaxDiscToSU") 
	End Property
	
	Public Property Get CartType
		CartType = GetDBConfigValue("CartType") 
	End Property
	
	Public Property Get UseCustomTransMsg
		UseCustomTransMsg = GetDBConfigValue("UseCustomTransMsg") 
	End Property
	
	Public Property Get CustomTransMsg
		CustomTransMsg = GetDBConfigValue("CustomTransMsg") 
	End Property
	
	Public Property Get ShowLineDiscount
		ShowLineDiscount = GetDBConfigValue("ShowLineDiscount") 
	End Property
	
	Public Property Get PrintPriceBefDiscount
		PrintPriceBefDiscount = GetDBConfigValue("PrintPriceBefDiscount") 
	End Property
	
	Public Property Get PrintLineDiscount
		PrintLineDiscount = GetDBConfigValue("PrintLineDiscount") 
	End Property
	
	Public Property Get AllowClientPartSuppSel
		AllowClientPartSuppSel = GetDBConfigValue("AllowClientPartSuppSel") 
	End Property
	
	Public Property Get EnSelAll
		EnSelAll = GetDBConfigValue("EnSelAll") 
	End Property
	
	Public Property Get EnSellAllUnitFrom
		EnSellAllUnitFrom = GetDBConfigValue("EnSellAllUnitFrom") 
	End Property

	'Anon Login
	Public Property Get EnableAnSesion
		EnableAnSesion = GetDBConfigValue("EnableAnSesion") 
	End Property
	
	Public Property Get EnableAnReg
		EnableAnReg = GetDBConfigValue("EnableAnReg") 
	End Property
	
	Public Property Get WebAddress
		WebAddress = GetDBConfigValue("WebAddress") 
	End Property
	
	Public Property Get AnSesListNum
		AnSesListNum = GetDBConfigValue("AnSesListNum") 
	End Property
	
	Public Property Get AnonSesFilter
		AnonSesFilter = GetDBConfigValue("AnonSesFilter") 
	End Property
	
	Public Property Get EnableAnRegTerms
		EnableAnRegTerms = GetDBConfigValue("EnableAnRegTerms") 
	End Property
	
	Public Property Get AnTerms
		AnTerms = GetDBConfigValue("AnTerms") 
	End Property
	
	Public Property Get TopLogo
		TopLogo = GetDBConfigValue("TopLogo") 
	End Property
	
	Public Property Get MailLogo
		MailLogo = GetDBConfigValue("MailLogo") 
	End Property
	
	Public Property Get AgentLogo
		AgentLogo = GetDBConfigValue("AgentLogo") 
	End Property
	
	Public Property Get AnRegConfAsignSLP
		AnRegConfAsignSLP = GetDBConfigValue("AnRegConfAsignSLP") 
	End Property
	
	Public Property Get EnChooseCType
		EnChooseCType = GetDBConfigValue("EnChooseCType") 
	End Property
	
	Public Property Get AnRegConfFrom
		AnRegConfFrom = GetDBConfigValue("AnRegConfFrom") 
	End Property
	
	Public Property Get AnRegConfTo
		AnRegConfTo = GetDBConfigValue("AnRegConfTo") 
	End Property
	
	Public Property Get ClientType
		ClientType = GetDBConfigValue("ClientType") 
	End Property
	
	Public Property Get AnRegAct
		AnRegAct = GetDBConfigValue("AnRegAct") 
	End Property
	
	Public Property Get RegActMailAdd
		RegActMailAdd = GetDBConfigValue("RegActMailAdd") 
	End Property
		
	Public Property Get RemPwdMailAdd
		RemPwdMailAdd = GetDBConfigValue("RemPwdMailAdd") 
	End Property
	
	Public Property Get AnRegConfRejNote
		AnRegConfRejNote = GetDBConfigValue("AnRegConfRejNote") 
	End Property
	
	'Admin Client Settings
	
	Public Property Get EnableDROnlyNote
		EnableDROnlyNote = GetDBConfigValue("EnableDROnlyNote") 
	End Property
	
	Public Property Get MyDataReadOnly
		MyDataReadOnly = GetDBConfigValue("MyDataReadOnly") 
	End Property
	
	'Admin Auto Gen
	
	Public Property Get AutoGenOCRD
		AutoGenOCRD = GetDBConfigValue("AutoGenOCRD")
	End Property
	
	Public Property Get AutoGenOITM
		AutoGenOITM = GetDBConfigValue("AutoGenOITM")
	End Property
	
	'Verfy Orders
	
	Public Property Get VerfyBtchOrder
		VerfyBtchOrder = GetDBConfigValue("VerfyBtchOrder") 
	End Property
	
	Public Property Get Verfy3dxOrder
		Verfy3dxOrder = GetDBConfigValue("Verfy3dxOrder") 
	End Property
	
	'Admin Doc Conf
	
	Public Property Get ClientReservedInvoice
		ClientReservedInvoice = GetDBConfigValue("ClientReservedInvoice") 
	End Property
	
	Public Property Get EnResInv
		EnResInv = GetDBConfigValue("EnResInv") 
	End Property
	
	Public Property Get DefResInv
		DefResInv = GetDBConfigValue("DefResInv") 
	End Property
		
	Public Property Get ORCTContraComp
		ORCTContraComp = GetDBConfigValue("ORCTContraComp") 
	End Property
	
	Public Property Get ApplyOpenRctToInvBal
		ApplyOpenRctToInvBal = GetDBConfigValue("ApplyOpenRctToInvBal") 
	End Property
	
	Public Property Get ChecksFilter
		ChecksFilter = GetDBConfigValue("ChecksFilter") 
	End Property
	
	Public Property Get IgnoreSystemChecksFilter
		IgnoreSystemChecksFilter = GetDBConfigValue("IgnoreSystemChecksFilter") 
	End Property
	
	'Admin Obj Conf
	
	Public Property Get EnableOSCL
		EnableOSCL = GetDBConfigValue("ActiveObjectA191") and myAut.HasAuthorization(189)
	End Property
	
	Public Property Get EnableOCTR
		EnableOCTR = GetDBConfigValue("ActiveObjectA190") and myAut.HasAuthorization(191)
	End Property
	
	Public Property Get EnableOINS
		EnableOINS = GetDBConfigValue("ActiveObjectA176") and myAut.HasAuthorization(190)
	End Property
	
	Public Property Get EnableOOPR
		EnableOOPR = GetDBConfigValue("ActiveObjectA97") and myAut.HasAuthorization(178)
	End Property
	
	Public Property Get EnableOWTQ
		EnableOWTQ = GetDBConfigValue("ActiveObjectA1250000001") and myAut.HasAuthorization(192)
	End Property
	
	Public Property Get EnableOWTR
		EnableOWTR = GetDBConfigValue("ActiveObjectA67") and myAut.HasAuthorization(193)
	End Property
	
	Public Property Get EnableOCLG
		EnableOCLG = GetDBConfigValue("ActiveObjectA33") and myAut.HasAuthorization(67)
	End Property
	
	Public Property Get EnableOCRD
		EnableOCRD = GetDBConfigValue("ActiveObjectA2") and myAut.HasBPCreateAccess
	End Property
	
	Public Property Get EnableOITM 
		EnableOITM = GetDBConfigValue("ActiveObjectA4") and myAut.HasAuthorization(44)
	End Property
	
	Public Property Get EnableOQUT 
		Select Case userType
			Case "V"
				EnableOQUT = GetDBConfigValue("ActiveObjectA23") and myAut.HasAuthorization(30)
			Case "C"
				EnableOQUT = GetDBConfigValue("ActiveObjectC23")
		End Select
	End Property
	
	Public Property Get EnableORDR
		Select Case userType
			Case "V"
				EnableORDR = GetDBConfigValue("ActiveObjectA17") and myAut.HasAuthorization(31)
			Case "C"
				EnableORDR = GetDBConfigValue("ActiveObjectC17")
		End Select
	End Property
	
	Public Property Get EnableODPIReq
		EnableODPIReq = GetDBConfigValue("ActiveObjectA203") and myAut.HasAuthorization(176)
	End Property
	
	Public Property Get EnableODPIInv
		EnableODPIInv = GetDBConfigValue("ActiveObjectA204") and myAut.HasAuthorization(177)
	End Property
	
	Public Property Get EnableOPOR
		EnableOPOR = GetDBConfigValue("ActiveObjectA22") and myAut.HasAuthorization(82)
	End Property
	
	Public Property Get EnableOPQT
		EnableOPQT = GetDBConfigValue("ActiveObjectA540000006") and myAut.HasAuthorization(181)
	End Property
	
	Public Property Get EnableOINV 
		EnableOINV = GetDBConfigValue("ActiveObjectA13") and myAut.HasAuthorization(34)
	End Property
	
	Public Property Get EnableOINVRes 
		EnableOINVRes = GetDBConfigValue("ActiveObjectA-13") and myAut.HasAuthorization(171)
	End Property
	
	Public Property Get EnableORCT 
		EnableORCT = GetDBConfigValue("ActiveObjectA24") and myAut.HasAuthorization(27)
	End Property
	
	Public Property Get EnableODLN 
		Select Case userType
			Case "V"
				EnableODLN = GetDBConfigValue("ActiveObjectA15") and myAut.HasAuthorization(29)
			Case "C"
				EnableODLN = GetDBConfigValue("ActiveObjectC15")
		End Select
	End Property
	
	Public Property Get EnableCashInv 
		Select Case userType
			Case "V"
				EnableCashInv = GetDBConfigValue("ActiveObjectA48") and myAut.HasAuthorization(35)
			Case "C"
				EnableCashInv = GetDBConfigValue("ActiveObjectC48")
		End Select
	End Property

	'DB Settings
	Function GetEnableMinInv
		Select Case userType
			Case "C"
				GetEnableMinInv = EnableMinInv
			Case "V"
				GetEnableMinInv = EnableMinInvV
		End Select
	End Function
	
	Function GetMinInv
		Select Case userType
			Case "C"
				GetMinInv = MinInv
			Case "V"
				GetMinInv = MinInvV
		End Select
	End Function
	
	Function GetMinInvBy
		Select Case userType
			Case "C"
				GetMinInvBy = MinInvBy
			Case "V"
				GetMinInvBy = MinInvVBy
		End Select
	End Function
	
	Function GetShowImg
		Select Case userType
			Case "C"
				GetShowImg = ShowClientImg
			Case "V"
				GetShowImg = ShowAgentImg
		End Select
	End Function

	Function GetApplyGenFilter
		If Not IsNull(GenFilter) Then
			Select Case userType
				Case "C"
					GetApplyGenFilter = GenFAppC
				Case "V"
					GetApplyGenFilter = GenFAppV
			End Select
		Else
			GetApplyGenFilter = False
		End If
	End Function	
	
	Function GetGenFilter
		GetGenFilter = Replace(Replace(Replace(GenFilter, "@SlpCode", Session("vendid")), "@CardCode", "N'" & saveHTMLDecode(Session("UserName"), False) & "'"), "@UserType", "'" & userType & "'")
	End Function
	
	Function GetDefCatOrdr
		Select Case userType
			Case "C"
				GetDefCatOrdr = DefCatOrdrC
			Case "V"
				GetDefCatOrdr = DefCatOrdrV
		End Select
	End Function

	Function GetShowRef
		Select Case userType
			Case "C"
				GetShowRef = ShowClientRef
			Case "V"
				GetShowRef = True
		End Select
	End Function

	Function GetShowSalUn
		Select Case userType
			Case "C"
				GetShowSalUn = ShowClientSalUn
			Case "V"
				GetShowSalUn = EnableUnitSelection
		End Select
	End Function

	Function GetSaleUnit
		Select Case userType
			Case "C"
				GetSaleUnit = ClientSaleUnit
			Case "V"
				GetSaleUnit = AgentSaleUnit
		End Select
	End Function
	
	Function GetDefView
		Select Case userType
			Case "C"
				GetDefView = DefViewCL
			Case "V"
				GetDefView = DefViewAG
		End Select
	End Function 
	
	Function GetAfterCartAdd
		Select Case userType
			Case "C"
				GetAfterCartAdd = AfterCartAddC
			Case "V"
				GetAfterCartAdd = AfterCartAddV
		End Select
	End Function 
	
	Function GetEnableCartImp
		Select Case userType
			Case "C"
				GetEnableCartImp = EnableCartImpC
			Case "V"
				GetEnableCartImp = EnableCartImpV
		End Select
	End Function 
	
	Function GetShowCxcOpenInv 
		Select Case userType
			Case "C"
				GetShowCxcOpenInv = showCxcOpenInvC
			Case "V"
				GetShowCxcOpenInv = showCxcOpenInv
		End Select
	End Function 
	
	Function GetShowCxcOpenInvBy  
		Select Case userType
			Case "C"
				GetShowCxcOpenInvBy = showCxcOpenInvByC
			Case "V"
				GetShowCxcOpenInvBy = showCxcOpenInvBy
		End Select
	End Function 
	
	Function GetShowCxcIncTrans   
		Select Case userType
			Case "C"
				GetShowCxcIncTrans = showCxcIncTransC
			Case "V"
				GetShowCxcIncTrans = showCxcIncTrans
		End Select
	End Function
	
	Function GetShowCxcDueDate 
		Select Case userType
			Case "C"
				GetShowCxcDueDate = showCxcDueDateC
			Case "V"
				GetShowCxcDueDate = showCxcDueDate
		End Select
	End Function 
	
	Function GetShowQtyInUn
		Select Case userType
			Case "C"
				GetShowQtyInUn = ShowQtyInUnCl
			Case "V"
				GetShowQtyInUn = ShowQtyInUnAg
		End Select
	End Function 
	
	Public Function ConcValue(ByVal Value, ByVal AddValue) 
		If Value <> "" Then Value = Value & ", " 
		Value = Value & AddValue 
		ConcValue = Value 
	End Function 
	
	Public Sub WriteDebug(ByVal str)
		Response.Write "<textarea id=""taDebug"" style=""width: 100%; height:200px;"">" & Server.HTMLEncode(str) & "</textarea>"
	End Sub
End Class
	

%>
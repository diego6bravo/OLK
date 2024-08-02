<!--#include file="lang/adminMenus.asp" -->
<%

curPage = Right(Request.ServerVariables("URL"),Len(Request.ServerVariables("URL"))-InStrRev(Request.ServerVariables("URL"),"/"))

'adminShipSis.asp|" & getadminMenusLngStr("LmnuShipSis") & ", adminClientsPriceList.asp|" & getadminMenusLngStr("LmnuClientPList") & ", 
mnuGen = "adminGeneral.asp|" & getadminMenusLngStr("LmnuGenOpt") & ", adminCatProp.asp|" & getadminMenusLngStr("LmnuCatPropOpt") & ", adminCart.asp|" & getadminMenusLngStr("LmnuShopCart") & ", adminInv.asp|" & getadminMenusLngStr("LmnuInv") & ", adminBranchs.asp|" & getadminMenusLngStr("LmnuBranches") & ", adminDocConf.asp|" & getadminMenusLngStr("LmnuObjsConf") & ", adminPaySis.asp|" & getadminMenusLngStr("LmnuPaySis") & ", adminAlerts.asp|" & getadminMenusLngStr("LmnuAlerts") & ", adminUsersLng.asp|" & getadminMenusLngStr("LmnuUsersLng") & "" 'adminAgentsPriceList.asp|" & getadminMenusLngStr("LmnuAgentPList") & ", 

'adminMsg.asp|" & getadminMenusLngStr("LmnuEmails") & ", 
mnuEnvir = "adminLanguages.asp|" & getadminMenusLngStr("LmnuLanguages") & ", adminLogos.asp|" & getadminMenusLngStr("LmnuLogo") & ", adminDec.asp|" & getadminMenusLngStr("LmnuClientDes") & ", adminAnonLogin.asp|" & getadminMenusLngStr("LmnuAnonSes") & ", adminSec.asp?UType=C|" & getadminMenusLngStr("LmnuSec") & " " & getadminMenusLngStr("DtxtClients") & ", " & _
"adminSec.asp?UType=A|" & getadminMenusLngStr("DtxtForms") & " " & getadminMenusLngStr("DtxtAgents") & ", adminSec.asp?UType=P|" & getadminMenusLngStr("DtxtForms") & " " & getadminMenusLngStr("DtxtPocket") & ", adminBN.asp|" & getadminMenusLngStr("LmnuBanners") & ", adminNews.asp|" & getadminMenusLngStr("LmnuNews") & ", adminPolls.asp|" & getadminMenusLngStr("LtxtPolls") & ""

mnuAcc = "adminPwd.asp|" & getadminMenusLngStr("LmnuPwd") & ", adminClientsAccess.asp|" & getadminMenusLngStr("LmnuClientsAccess") & ""

If Not myApp.SingleSignOn Then 
	mnuAcc = mnuAcc & ", adminAgentsAccess.asp|" & getadminMenusLngStr("LmnuAgentsAccess") & ""
Else
	mnuAcc = mnuAcc & ", adminSingleAccess.asp|" & getadminMenusLngStr("LmnuAgentsAccess") & ""
End If

mnuAcc = mnuAcc & ", adminAutGrp.asp|" & getadminMenusLngStr("LmnuAutGrp") & ", adminDocFlow.asp|" & getadminMenusLngStr("LmnuDocFlow") & ", adminCustomLogin.asp|" & getadminMenusLngStr("LmnuCustLogin") & ""

'", adminAut.asp|" & getadminMenusLngStr("LtxtAut") & ""

If Session("olkdb") = "" Then 
	mnuAcc = "adminPwd.asp|" & getadminMenusLngStr("LmnuPwd") & ""
	If myApp.SingleSignOn Then mnuAcc = mnuAcc & ", adminSingleAccess.asp|" & getadminMenusLngStr("LmnuAgentsAccess") & ""
End If

mnuPers = "adminCustDec.asp|" & getadminMenusLngStr("LmnuCustDec") & ", adminInformer.asp|" & getadminMenusLngStr("LmnuInformer") & ", adminPrintTitle.asp|" & getadminMenusLngStr("LmnuPrintTitle") & ", adminCardOpt.asp|" & getadminMenusLngStr("LmnuCardDet") & ", adminInvOpt.asp|" & getadminMenusLngStr("LmnuItemDet") & ", adminiPO.asp|" & getadminMenusLngStr("LmnuiPO") & ", adminBatchOpt.asp|" & getadminMenusLngStr("LmnuBtchDet") & ", adminCartOpt.asp|" & getadminMenusLngStr("LmnuCartDet") & ", adminCUFD.asp|" & getadminMenusLngStr("LmnuUDF") & ", adminDocBreakDown.asp|" & getadminMenusLngStr("LmnuBreakDown") & ", adminAlterNames.asp|" & getadminMenusLngStr("LmnuAlterNames") & ", adminCatOpt.asp|" & getadminMenusLngStr("LmnuCustCat") & ", adminMenuGroups.asp|" & getadminMenusLngStr("LmnuSearchTree") & ", adminCustomSearch.asp?ObjID=2|" & getadminMenusLngStr("LmnuCustomSearchCl") & ", adminCustomSearch.asp?ObjID=4|" & getadminMenusLngStr("LmnuCustomSearch") & ", adminCartMinRep.asp|" & getadminMenusLngStr("LmnuCartMiniRep") & ", adminObjConfCols.asp|" & getadminMenusLngStr("LmnuObjConfCols") & ", adminObjPrint.asp|" & getadminMenusLngStr("LmnuPrint") & ", adminPriceCod.asp|" & getadminMenusLngStr("LmnuNumCod") & ", adminDefObjs.asp|" & getadminMenusLngStr("LmnuCustObjs") & ", adminFooter.asp|" & getadminMenusLngStr("LmnuFooter") & ", adminLayout.asp|" & getadminMenusLngStr("LmnuLayout") & ", adminSmallCat.asp|" & getadminMenusLngStr("LmnuSmallCat") & ", adminCatNav.asp|" & getadminMenusLngStr("LmnuCatNav") & ", adminOps.asp|" & getadminMenusLngStr("LtxtOperations") & ""

mnuRep = "adminReps.asp?uType=C|" & getadminMenusLngStr("LmnuClientReps") & ", adminReps.asp?uType=V|" & getadminMenusLngStr("LmnuAgentsRep") & ""

If mnuAgentsRep = "" then
	mnuGen 		= Replace(mnuGen, "||", "|")
	mnuEnvir 	= Replace(mnuEnvir, "||", "|")
	mnuAcc 		= Replace(mnuAcc, "||", "|")
	mnuPers 	= Replace(mnuPers, "||", "|")
	mnuRep 		= Replace(mnuRep, "||", "|")
End If

Select Case curPage
	'Inicio
	Case "admin.asp"
		CurSection = "" & getadminMenusLngStr("LmnuHomeOpt") & ""
	'Generales	
	Case "adminGeneral.asp"
		CurSection = "" & getadminMenusLngStr("LmnuGenOpt") & ""
		ShowGeneral = True	
	Case "adminCatProp.asp"
		CurSection = "" & getadminMenusLngStr("LmnuCatPropOpt") & ""
		ShowGeneral = True
	Case "adminCart.asp"
		CurSection = "" & getadminMenusLngStr("LmnuShopCart") & ""
		ShowGeneral = True					
	Case "adminInv.asp"
		CurSection = "" & getadminMenusLngStr("LmnuInv") & ""
		ShowGeneral = True			
	Case "adminBranchs.asp"
		CurSection = "" & getadminMenusLngStr("LmnuBranches") & ""
		ShowGeneral = True	
	Case "adminBranchsEdit.asp"
		CurSection = "" & getadminMenusLngStr("LmnuEditBranch") & ""
		ShowGeneral = True	
	Case "adminDocConf.asp"
		CurSection = "" & getadminMenusLngStr("LmnuObjsConf") & ""
		ShowGeneral = True		
	Case "adminPaySis.asp"
		CurSection = "" & getadminMenusLngStr("LmnuPaySis") & ""
		ShowGeneral = True	
	Case "adminShipSis.asp"
		CurSection = "" & getadminMenusLngStr("LmnuShipSis") & ""
		ShowGeneral = True
	Case "adminClientsPriceList.asp"
		CurSection = "" & getadminMenusLngStr("LmnuClientPList") & ""
		ShowGeneral = True		
	Case "adminAgentsPriceList.asp"
		CurSection = "" & getadminMenusLngStr("LmnuAgentPList") & ""
		ShowGeneral = True		
	Case "adminAlerts.asp"
		CurSection = "" & getadminMenusLngStr("LmnuAlerts") & ""
		ShowGeneral = True
	Case "adminUsersLng.asp"
		CurSection = "" & getadminMenusLngStr("LmnuUsersLng") & ""
		ShowGeneral = True
		
	'Ambiente
	Case "adminLogos.asp"
		CurSection = "" & getadminMenusLngStr("LmnuLogo") & ""
		ShowEnvir = True
	Case "adminCartMore.asp"
		CurSection = getadminMenusLngStr("DtxtCart") & " - " & getadminMenusLngStr("LtxtMoreOpt") & ""
		ShowEnvir = True
	Case "adminDec.asp"
		CurSection = "" & getadminMenusLngStr("LmnuClientDes") & ""
		ShowEnvir = True	
	Case "adminLanguages.asp"
		CurSection = "" & getadminMenusLngStr("LmnuLanguages") & ""
		ShowEnvir = True	
	Case "adminAnonLogin.asp"
		CurSection = "" & getadminMenusLngStr("LmnuAnonSes") & ""
		ShowEnvir = True		
	Case "adminSec.asp"
		curPage = curPage & "?UType=" & Request("UType")
		Select Case Request("UType")
			Case "C"
				CurSection = "" & getadminMenusLngStr("LmnuSec") & " " & getadminMenusLngStr("DtxtClients")
			Case "P"
				CurSection = getadminMenusLngStr("DtxtForms") & " " & getadminMenusLngStr("DtxtPocket")
			Case "A"
				CurSection = getadminMenusLngStr("DtxtForms") & " " & getadminMenusLngStr("DtxtAgents")
		End Select
		ShowEnvir = True
	Case "adminSecIndex.asp"
		CurSection = "" & getadminMenusLngStr("LmnuSecIndex") & ""
		ShowEnvir = True
	Case "adminMyData.asp"
		CurSection = "" & getadminMenusLngStr("LmnuMyData") & ""
		ShowEnvir = True
	Case "adminSecEdit.asp"
		If Request("SecID") = "" Then
			Select Case Request("UType")
				Case "C"
					CurSection = "" & getadminMenusLngStr("LmnuAddSec") & "" 
				Case "P", "A"
					CurSection = "" & getadminMenusLngStr("LmnuAddForm") & ""
			End Select
		Else 
			Select Case Request("UType")
				Case "C"
					CurSection = "" & getadminMenusLngStr("LmnuEditSec") & ""
				Case "P", "A"
					CurSection = "" & getadminMenusLngStr("LmnuEditForm") & ""
			End Select
		End If
		ShowEnvir = True
	Case "adminMsg.asp"
		CurSection = "" & getadminMenusLngStr("LmnuEmails") & ""
		ShowEnvir = True
		If Request("MsgID") <> "" Then mnuShowNormal = True
	Case "adminBN.asp"
		CurSection = "" & getadminMenusLngStr("LmnuBanners") & ""
		ShowEnvir = True
	Case "adminBNEdit.asp"
		If Request("BannerID") = "" Then CurSection = "" & getadminMenusLngStr("LmnuNewBanner") & "" Else CurSection = "" & getadminMenusLngStr("LmnuEditBanner") & ""
		ShowEnvir = True
	Case "adminNews.asp"
		CurSection = "" & getadminMenusLngStr("LmnuNews") & ""
		ShowEnvir = True
	Case "adminPolls.asp"
		CurSection = "" & getadminMenusLngStr("LtxtPolls") & ""
		ShowEnvir = True
	Case "adminNewsEdit.asp"
		If Request("newsIndex") = "" Then CurSection = "" & getadminMenusLngStr("LmnuNewNews") & "" Else CurSection = "" & getadminMenusLngStr("LmnuEditNews") & ""
		ShowEnvir = True
	Case "adminPollEdit.asp"
		If Request("pollIndex") = "" Then CurSection = "" & getadminMenusLngStr("LmnuNewPoll") & "" Else CurSection = "" & getadminMenusLngStr("LmnuEditPoll") & ""
		ShowEnvir = True
		
	'Accesos
	Case "adminCustomLogin.asp"
		CurSection = "" & getadminMenusLngStr("LmnuCustLogin") & ""
		ShowAcc = True
	Case "adminPwd.asp"
		CurSection = "" & getadminMenusLngStr("LmnuPwd") & ""
		ShowAcc = True
	Case "adminClientsAccess.asp"
		CurSection = "" & getadminMenusLngStr("LmnuClientsAccess") & ""
		ShowAcc = True		
	Case "adminAgentsAccess.asp"
		CurSection = "" & getadminMenusLngStr("LmnuAgentsAccess") & ""
		ShowAcc = True	
	Case "adminSingleAccess.asp"
		CurSection = "" & getadminMenusLngStr("LmnuAgentsAccess") & ""
		ShowAcc = True
	Case "adminSingleAccessDB.asp"
		CurSection = "" & getadminMenusLngStr("LmnuAgentsAccess") & ""
		ShowAcc = True	
	Case "adminDocFlow.asp"
		CurSection = "" & getadminMenusLngStr("LmnuDocFlow") & ""
		ShowAcc = True		
		If Request("FlowID") <> "" or Request("NewFlow") = "Y" Then mnuShowNormal = True
	Case "adminAut.asp"
		CurSection = "" & getadminMenusLngStr("LtxtAut") & ""
		ShowAcc = True
	Case "adminAutGrp.asp"
		Select Case Request("GrpID")
			Case "New"
				CurSection = "" & getadminMenusLngStr("LmnuAddAutGrp") & ""
				mnuShowNormal = True
			Case ""
				CurSection = "" & getadminMenusLngStr("LmnuAutGrp") & ""
			Case Else
				CurSection = "" & getadminMenusLngStr("LmnuEditAutGrp") & ""
				mnuShowNormal = True
		End Select
		ShowAcc = True		
	'Personalizaci�n
	Case "adminOpsEdit.asp"
		CurSection = "" & getadminMenusLngStr("LtxtOperations") & ""
		If CInt(Request("ID")) <> -1 Then
			CurSection = CurSection & " - " & getadminMenusLngStr("LmnuEditOp") & ""
		Else
			CurSection = CurSection & " - " & getadminMenusLngStr("LmnuNewOp") & ""
		End If
		ShowPer = True
	Case "adminOps.asp"
		CurSection = "" & getadminMenusLngStr("LtxtOperations") & ""
		ShowPer = True
	Case "adminCustDec.asp"
		CurSection = "" & getadminMenusLngStr("LmnuCustDec") & ""
		ShowPer = True
	Case "adminObjPrint.asp"
		CurSection = "" & getadminMenusLngStr("LmnuPrint") & ""
		ShowPer = True
	Case "adminInformer.asp"
		CurSection = "" & getadminMenusLngStr("LmnuInformer") & ""
		ShowPer = True
	Case "adminInformerEdit.asp"
		If Request("ID") <> "" Then
			CurSection = "" & getadminMenusLngStr("LttlEditMonitor") & ""
		Else
			CurSection = "" & getadminMenusLngStr("LttlAddMonitor") & ""
		End If
		ShowPer = True
	Case "adminCustomSearch.asp"
		CurPage = CurPage & "?ObjID=" & Request("ObjID")
		Select Case CInt(Request("ObjID"))
			Case 2
				CurSection = "" & getadminMenusLngStr("LmnuCustomSearchCl") & ""
			Case 4
				CurSection = "" & getadminMenusLngStr("LmnuCustomSearch") & ""
		End Select
		ShowPer = True
	Case "adminCustomSearchEdit.asp"
		Select Case CInt(Request("ObjID"))
			Case 2
				CurSection = "" & getadminMenusLngStr("LmnuCustomSearchCl") & ""
			Case 4
				CurSection = "" & getadminMenusLngStr("LmnuCustomSearch") & ""
		End Select
		If Request("ID") <> "" Then
			CurSection = CurSection & " - " & getadminMenusLngStr("LmnuEditSearch") & ""
		Else
			CurSection = CurSection & " - " & getadminMenusLngStr("LmnuNewSearch") & ""
		End If
		ShowPer = True
	Case "adminCardOpt.asp"
		CurSection = "" & getadminMenusLngStr("LmnuCardDet") & ""
		ShowPer = True
		If Request("edit") = "Y" or Request("NewFld") = "Y" Then mnuShowNormal = True
	Case "adminPrintTitle.asp"
		CurSection = "" & getadminMenusLngStr("LmnuPrintTitle") & ""
		ShowPer = True
		If Request("edit") = "Y" or Request("NewFld") = "Y" Then mnuShowNormal = True
	Case "adminBatchOpt.asp"
		CurSection = "" & getadminMenusLngStr("LmnuBtchDet") & ""
		If Request("edit") = "Y" Then
			CurSection = CurSection & " - " & getadminMenusLngStr("LmnuEditField") & ""
		ElseIf Request("NewFld") = "Y" Then
			CurSection = CurSection & " - " & getadminMenusLngStr("LmnuNewFld") & ""
		End If
		ShowPer = True
		If Request("edit") = "Y" or Request("NewFld") = "Y" Then mnuShowNormal = True
	Case "adminCartOpt.asp"
		CurSection = "" & getadminMenusLngStr("LmnuCartDet") & ""
		If Request("edit") = "Y" Then
			CurSection = CurSection & " - " & getadminMenusLngStr("LmnuEditField") & ""
		ElseIf Request("NewFld") = "Y" Then
			CurSection = CurSection & " - " & getadminMenusLngStr("LmnuNewFld") & ""
		End If
		ShowPer = True
		If Request("edit") = "Y" or Request("NewFld") = "Y" Then mnuShowNormal = True
	Case "adminAlterNames.asp"
		CurSection = "" & getadminMenusLngStr("LmnuAlterNames") & ""
		ShowPer = True
		If Request("AlterLng") <> "" Then mnuShowNormal = True
	Case "adminInvOpt.asp"
		CurSection = "" & getadminMenusLngStr("LmnuItemDet") & ""
		If Request("edit") = "Y" Then
			CurSection = CurSection & " - " & getadminMenusLngStr("LmnuEditField") & ""
		ElseIf Request("NewFld") = "Y" Then
			CurSection = CurSection & " - " & getadminMenusLngStr("LmnuNewFld") & ""
		End If
		ShowPer = True
		If Request("edit") = "Y" or Request("NewFld") = "Y" Then mnuShowNormal = True
	Case "adminiPO.asp"
		CurSection = "" & getadminMenusLngStr("LmnuiPO") & ""
		If Request("edit") = "Y" Then
			CurSection = CurSection & " - " & getadminMenusLngStr("LmnuEditField") & ""
		ElseIf Request("NewFld") = "Y" Then
			CurSection = CurSection & " - " & getadminMenusLngStr("LmnuNewFld") & ""
		End If
		ShowPer = True
		If Request("edit") = "Y" or Request("NewFld") = "Y" Then mnuShowNormal = True
	Case "adminCUFD.asp"
		CurSection = "" & getadminMenusLngStr("LmnuUDF") & ""
		ShowPer = True
		If Request("sType") <> "" Then mnuShowNormal = True
	Case "adminDocBreakDown.asp"
		CurSection = "" & getadminMenusLngStr("LmnuBreakDown") & ""
		ShowPer = True
	Case "adminMenuGroups.asp"
		CurSection = "" & getadminMenusLngStr("LmnuSearchTree") & ""
		ShowPer = True
		If Request("editID") <> "" or Request("new") = "Y" Then mnuShowNormal = True
	Case "adminCatOpt.asp"
		CurSection = "" & getadminMenusLngStr("LmnuCustCat") & ""
		If Request("edit") = "Y" Then
			CurSection = CurSection & " - " & getadminMenusLngStr("LmnuEditField") & ""
		ElseIf Request("NewFld") = "Y" Then
			CurSection = CurSection & " - " & getadminMenusLngStr("LmnuNewFld") & ""
		End If
		ShowPer = True
		If Request("edit") = "Y" or Request("NewFld") = "Y" Then mnuShowNormal = True
	Case "adminCartMinRep.asp"
		CurSection = "" & getadminMenusLngStr("LmnuCartMiniRep") & ""
		ShowPer = True
	Case "adminCartMinRepEdit.asp"
		CurSection = "" & getadminMenusLngStr("LmnuCartMiniRep") & ""
		If Request("goAction") = "editLine" or Request("btnApply") <> "" or Request("btnRestore") <> "" Then
			CurSection = CurSection & " - " & getadminMenusLngStr("LmnuEditField") & ""
		Else
			CurSection = CurSection & " - " & getadminMenusLngStr("LmnuNewFld") & ""
		End If
		ShowPer = True
	Case "adminObjConfCols.asp"
		CurSection = "" & getadminMenusLngStr("LmnuObjConfCols") & ""
		If Request("ID") <> "" Then
			CurSection = CurSection & " - " & getadminMenusLngStr("LmnuEditField") & ""
		ElseIf Request("New") = "Y" Then
			CurSection = CurSection & " - " & getadminMenusLngStr("LmnuNewFld") & ""
		End If
		ShowPer = True
	Case "adminPriceCod.asp"
		CurSection = "" & getadminMenusLngStr("LmnuNumCod") & ""
		ShowPer = True
	Case "adminDefObjs.asp"
		CurSection = "" & getadminMenusLngStr("LmnuCustObjs") & ""
		ShowPer = True
	Case "adminFooter.asp"
		CurSection = "" & getadminMenusLngStr("LmnuFooter") & ""
		ShowPer = True
	Case "adminLayout.asp"
		CurSection = "" & getadminMenusLngStr("LmnuLayout") & ""
		ShowPer = True
	Case "adminSmallCat.asp"
		CurSection = "" & getadminMenusLngStr("LmnuSmallCat") & ""
		ShowPer = True
	Case "adminDefObjEdit.asp"
		CurSection = "" & getadminMenusLngStr("LmnuPerObj") & ""
		ShowPer = True
	Case "adminCatNav.asp"
		CurSection = "" & getadminMenusLngStr("LmnuCatNav") & ""
		ShowPer = True
		If Request("editIndex") <> "" or Request("New") = "Y" Then mnuShowNormal = True
	'Reportes
	Case "adminReps.asp"
		curPage = curPage & "?uType=" & Request("uType")
		If Request("uType") = "V" Then
			CurSection = "" & getadminMenusLngStr("LtxtAgentsRep") & ""
		Else
			CurSection = "" & getadminMenusLngStr("LtxtClientsRep") & ""
		End If
		ShowRep = True
	Case "adminRepEdit.asp"
		CurSection = "" & getadminMenusLngStr("LmnuEditRep") & ""
		ShowRep = True
	Case "adminRepNew.asp"
		CurSection = "" & getadminMenusLngStr("LmnuNewRep") & ""
		ShowRep = True
	Case "adminReps.asp"
		CurSection = "" & getadminMenusLngStr("LmnuReps") & ""
		
	'Other
	Case "adminUpdate.asp"
		CurSection = "" & getadminMenusLngStr("LmnuUpdates") & ""
	Case "adminLicInf.asp"
		CurSection = "" & getadminMenusLngStr("LmnuLic") & ""
	Case "adminSystem.asp"
		CurSection = "" & getadminMenusLngStr("LtxtSysProp") & ""
End Select
%>
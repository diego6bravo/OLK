
<!--#include file="lang/agentTop.asp" -->
<% If session("OLKDB") = "" Then 
	response.redirect "lock.asp"
ElseIf userType = "C" Then
	response.redirect "default.asp"
ElseIf Session("OLKAdmin") Then
	response.redirect "admin/admin.asp"
End If

Response.Expires = 60
Response.Expiresabsolute = Now() - 1

If Request("cCode") <> "" Then Session("UserName") = saveHTMLDecode(Request("cCode"), True)
MainDoc = "agent.asp"

curCmd = CStr(Request("cmd"))

Select Case strScriptName
Case "clientssearch.asp"
	Menu = true
	SearchCmd = "clientsSearch"
Case "cart.asp"
	Menu = true
	SearchCmd = "searchCart"
	CartInfo = true
Case "ofertsman.asp"
	Menu = true
	SearchCmd = "searchOfertsX"
Case "searchopeneditems.asp"
	Menu = true
	SearchCmd = "searchItemX"
Case "searchopenedactivities.asp"
	Menu = true
	SearchCmd = "searchActX"
Case "searchopenedso.asp"
	Menu = true
	SearchCmd = "searchSOX"
Case "searchopenedcards.asp"
	Menu = true
	SearchCmd = "searchCardX"
Case "searchopeneddocs.asp"
	Menu = True
	SearchCmd = "searchDocX"
Case "activeclient.asp"
	Menu = True
	SearchCmd = "activeClient"
Case "search.asp"
	Menu = true
	If Session("RetVal") = "" Then 
		searchCmd = "searchCatalog" 
		pdfDoc = true
	Else 
		searchCmd = "searchCart"
	End If
	If Request("CPList") = "" and searchCmd = "searchCart" Then CartInfo = true
Case "adcustomsearch.asp"
	Menu = true
	Select Case CInt(Request("adObjID"))
		Case 2
			SearchCmd = "clientsSearch"
		Case 4
			If Session("RetVal") <> "" Then
				SearchCmd = "searchCart"
				CartInfo = true	
			Else
				SearchCmd = "searchCatalog"
			End If
		End Select
Case "wish.asp"
	Menu = true
	SearchCmd = "searchCart"
	CartInfo = true
Case "cxc.asp"
	If Session("RetVal") <> "" then
	Menu = true
	SearchCmd = "searchCart"
	End If
	excell = true
	pdfDoc = true
	pdfAction = "javascript:doCxcExport('pdf');"
	excellAction = "javascript:doCxcExport('excell')"
Case "cartsubmitconfirm.asp"
	pdfDoc  = true
Case "searchclient.asp"
	Session("Cart") = ""
	Session("RetVal") = ""
	Session("PayRetVal") = ""
	Session("PriceList") = ""
	Session("UserName") = ""
	SearchCmd = "searchClient"
Case "agentsearch.asp"
	Session("Cart") = ""
	Session("RetVal") = ""
	Session("PayRetVal") = ""
	Session("PriceList") = ""
	Session("UserName") = ""
Case "report.asp"
	Menu = true
	searchCmd = "report"
	excell = true
	excellAction = "javascript:saveRepPdf('Y');//"
	pdfDoc = true
	pdfAction = "javascript:saveRepPdf('N');//"
Case "executeconf.asp"
	searchCmd = "docsConfirmation"
	Menu = True
Case "extpollview.asp"
	searchCmd = "extPollview"
	Menu = True
End Select
If Not myAut.HasAuthorization(64) Then excell = false
pageVars = "?"
For each item in Request.Form
	PageVars = PageVars & item & "=" & Request(item)
	If i <> Request.Form.Count -1 Then PageVars = PageVars & "&"
next
For each item in Request.QueryString
	PageVars = PageVars & item & "=" & Request(item)
	If i <> Request.QueryString.Count -1 Then PageVars = PageVars & "&"
next

Select Case strScriptName
	Case "search.asp", "cxc.asp", "cartsubmitconfirm.asp", "rnews.asp", "finalpayment.asp", "report.asp", "newactivitysubmit.asp"
		EnablePrint = True
End Select
%>
<!--#include file="clearItem.asp"-->
<!--#include file="dateFormat.asp"-->
<% 

SelDes = 3
set rh = Server.CreateObject("ADODB.recordset") %>
<!--#include file="loadAlterNames.asp" -->
<!--#include file="repVars.inc" -->
<html <% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>Olk -&nbsp; <%=Replace(getagentTopLngStr("LtxtAgentInt"), "{0}", Server.HTMLEncode(txtAgents)) %></title>
<%
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKAgentTop" & Session("ID")
cmd.Parameters.Refresh()
cmd("@LanID") = Session("LanID")
If Session("username") <> "" and not IsNull(Session("username")) Then 
	cmd("@CardCode") = Session("UserName")
	ClientShowBalance = CheckAgentClientFilter(Session("UserName"), 2)
End If
cmd("@SlpCode") = Session("vendid")
cmd("@UserAccess") = Session("useraccess")
cmd("@UserType") = userType
cmd("@branch") = Session("branch")
If Request("rsIndex") <> "" Then cmd("@rsIndex") = Request("rsIndex")
If Session("UserAccess") = "U" Then cmd("@Sections") = myAut.AuthorizedForms 
set rh = cmd.execute()
EnableCartNewLine = rh("EnableCartNewLine") = "Y"
EnableClientActivation = myApp.AnRegAct = "C" or myApp.AnRegAct = "B"
optWish = rh("EnableSecWish") = "A" 
optOfert = rh("EnableSecOfer") = "A" 
CmpName = mySession.GetCompanyName
OLKVersion = rh("Version")
DirectRgate = rh("DirectRate")
EnableSections = rh("EnableSections") = "Y"
CardType = rh("CardType")
AgentName = mySession.GetAgentName
IsBPAssigned = rh("IsBPAssigned") = "Y" and not myAut.HasAuthorization(174) or myAut.HasAuthorization(174)
If Request("rsIndex") <> "" Then RSLastUpdate = rh("RSLastUpdate")
           
If Session("username") <> "" and not IsNull(Session("username")) Then
	ListNum = rh("OCRDListNum")
	CatalogFilter = rh("CatalogFilter")
	CatalogFilterAgent = rh("CatalogFilterAgent")
End If
		
If ListNum = -1 Then
	listNumAlrt = "" & getagentTopLngStr("LtxtValCLastPurListNu") & ""
ElseIf ListNum = -2 Then
	listNumAlrt = "" & getagentTopLngStr("LtxtValCLastDetListNu") & ""
End If
set rd = Server.CreateObject("ADODB.RecordSet")
sql = "select LanID from OLKDisLng"
rd.open sql, conn, 3, 1
myLng = ""
For i = 0 to UBound(myLanIndex)
	rd.Filter = "LanID = " & myLanIndex(i)(4)
	If rd.eof Then
		If myLng <> "" Then myLng = myLng & ", "
		myLng = myLng & myLanIndex(i)(0) & "{S}<span lang=""" & myLanIndex(i)(2) & """ dir=""" & myLanIndex(i)(3) & """>" & myLanIndex(i)(1) & "</span>"
	Else
		If CInt(myLanIndex(i)(4)) = CInt(Session("LanID")) Then
			Response.Redirect "agent.asp?newLng=" & myApp.NatLng
		End If
	End If
Next %>
<script language="javascript">
var txtValFldMaxChar = '<%=getagentTopLngStr("DtxtValFldMaxChar")%>';
var rtl = '<%=Session("rtl")%>';
var jsLangCol1 = '#0071E3';
var jsLangCol2 = '#EDF5FE';
var jsLangCol3 = '#DBEBFD';
var jsLangCol4 = '#FFFFFF';
var jsLangCol5 = '#4783C5';
var myLng = '<%=myLng%>';
var jsLangRev = false;
var strDateSep = '<%=Session("DateSep")%>';
<% If listNumAlrt <> "" Then %>function listNumAlrt() { alert('<%=listNumAlrt%>'); }<% End If %>
</script>
<script language="javascript" src="js_lang_sel.js"></script>
<script language="javascript" src="generalData.js.asp?dbID=<%=Session("ID")%>&LastUpdate=<%=myApp.LastUpdate%>"></script>
<script language="javascript" src="general.js"></script>
<script language="javascript" src="ventas.js"></script>
<script language="javascript" src="myMnu.js"></script>
<link type="text/css" href="design/0/jquery-ui-1.8.14.custom.css" rel="stylesheet" >	
<script type="text/javascript" src="jQuery/js/jquery-1.6.1.min.js"></script>
<script type="text/javascript" src="jQuery/js/jquery-ui-1.8.14.custom.min.js"></script>
<link rel="stylesheet" type="text/css" href="jQuery/css/ui.dropdownchecklist.themeroller.css">
<script type="text/javascript" src="jQuery/js/ui.dropdownchecklist-1.4-min.js"></script>

<% If strScriptName = "extpolledit.asp" Then %>
<script language="javascript">
var txtPredefined = '<%=getagentTopLngStr("LtxtPredefined")%>';
var txtCustomized = '<%=getagentTopLngStr("LtxtCustom")%>';
</script>
<link rel="stylesheet" href="js_color_picker_v2.css">
<script type="text/javascript" src="color_functions.js"></script>		
<script type="text/javascript" src="js_color_picker_v2.js"></script>
<% End If %>
<% If Request("rsIndex") <> "" Then %><link type="text/css" href="portal/viewRepCSS.asp?rsIndex=<%=Request("rsIndex")%>&LastUpdate=<%=RSLastUpdate%>" rel="stylesheet"><% End if %>

<script language="javascript">

<%
If myAut.HasAuthorization(1) Then
sql = "select T0.NavIndex, IsNull(AlterNavTitle, NavTitle) NavTitle, " & _
"Case When AutoRedir = 'Y' and (select Count('A') from OLKCatNavSub where NavIndex = T0.NavIndex) = 1 Then 'Y' Else 'N' End Redir, " & _
"(select top 1 SubIndex from OLKCatNavSub where NavIndex = T0.NavIndex) RedirIndex, " & _
"(select Case CatType When 'C' Then 'C' When 'S' Then 'T' End from OLKCatNav where NavIndex = (select top 1 SubIndex from OLKCatNavSub where NavIndex = T0.NavIndex)) RedirCatType " & _
"from OLKCatNav T0 " & _
"left outer join OLKCatNavAlterNames T1 on T1.NavIndex = T0.NavIndex and T1.LanID = " & Session("LanID") & " " & _
"where NavType = 'M' and Access in ('A', 'V') and Active = 'Y'  " & _
"and (DateDiff(day,ShowFrom,getdate()) >= 0 or ShowFrom is null) " & _
"and (DateDiff(day,getdate(),ShowTo) >= 0 or ShowTo is null) " & _
"order by T0.NavTitle "
set rs = conn.execute(sql) %>
var mnuCat = new Menu('Cat', false);
<% do while not rs.eof
If rs("Redir") = "N" Then
	navLink = "goLink(\'subNavIndex.asp?navIndex=" & rs("NavIndex") & "\');"
Else
	navLink = "goNavQry(" & rs("RedirIndex") & ", \'" & rs("RedirCatType") & "\')"
End If %>
mnuCat.addMenuItem(new MenuItem('<%=Replace(myHTMLEncode(rs("NavTitle")), "'", "\'")%>', '<%=navLink%>', false, null));
<% rs.movenext
loop
End If %>

var mnuNewDocs = new Menu('NewDocs', true);
<% 
comDocsMenu = myApp.EnableOPOR or myApp.EnableOQUT or myApp.EnableORDR or myApp.EnableODLN or myApp.EnableODPIReq or myApp.EnableODPIInv or myApp.EnableOINV or myApp.EnableOINVRes or  myApp.EnableCashInv
If myApp.EnableOPOR and CardType = "S" Then %>mnuNewDocs.addMenuItem(new MenuItem('<%=Replace(myHTMLEncode(txtOpor), "'", "\'")%>&nbsp;&nbsp;', 'createDocument(22);', false, null));<% End If %>
<% If myApp.EnableOQUT and CardType <> "S" Then %>mnuNewDocs.addMenuItem(new MenuItem('<%=Replace(myHTMLEncode(txtQuote), "'", "\'")%>&nbsp;&nbsp;', '<% If ListNum > -1 Then %>createDocument(23);<% Else %>listNumAlrt();<% End If %>', false, null));<% End If %>
<% If myApp.EnableORDR and CardType <> "S" Then %>mnuNewDocs.addMenuItem(new MenuItem('<%=Replace(myHTMLEncode(txtOrdr), "'", "\'")%>&nbsp;&nbsp;', '<% If ListNum > -1 Then %>createDocument(17);<% Else %>listNumAlrt();<% End If %>', false, null));<% End If %>
<% If CardType = "C" Then %>
<% If myApp.EnableODLN Then %>mnuNewDocs.addMenuItem(new MenuItem('<%=Replace(myHTMLEncode(txtOdln), "'", "\'")%>&nbsp;&nbsp;', '<% If ListNum > -1 Then %>createDocument(15);<% Else %>listNumAlrt();<% End If %>', false, null));<% End If %>
<% If myApp.EnableODPIReq Then %>mnuNewDocs.addMenuItem(new MenuItem('<%=Replace(myHTMLEncode(txtODPIReq), "'", "\'")%>&nbsp;&nbsp;', '<% If ListNum > -1 Then %>createDocument(203);<% Else %>listNumAlrt();<% End If %>', false, null));<% End If %>
<% If myApp.EnableODPIInv Then %>mnuNewDocs.addMenuItem(new MenuItem('<%=Replace(myHTMLEncode(txtODPIInv), "'", "\'")%>&nbsp;&nbsp;', '<% If ListNum > -1 Then %>createDocument(204);<% Else %>listNumAlrt();<% End If %>', false, null));<% End If %>
<% If myApp.EnableOINV Then %>mnuNewDocs.addMenuItem(new MenuItem('<%=Replace(myHTMLEncode(txtInv), "'", "\'")%>&nbsp;&nbsp;', '<% If ListNum > -1 Then %>createDocument(13);<% Else %>listNumAlrt();<% End If %>', false, null));<% End If %>
<% If myApp.EnableOINVRes Then %>mnuNewDocs.addMenuItem(new MenuItem('<%=Replace(myHTMLEncode(txtInvRes), "'", "\'")%>&nbsp;&nbsp;', '<% If ListNum > -1 Then %>createDocument(-13);<% Else %>listNumAlrt();<% End If %>', false, null));<% End If %>
<% If myApp.EnableCashInv Then %>mnuNewDocs.addMenuItem(new MenuItem('<%=Replace(myHTMLEncode(txtInv), "'", "\'")%>/<%=Replace(myHTMLEncode(txtRct), "'", "\'")%>&nbsp;&nbsp;', '<% If ListNum > -1 Then %>createDocument(48);<% Else %>listNumAlrt();<% End If %>', false, null));<% End If %>
<% End If %>

<% If EnableSections Then
SecErrMsg = "" & getagentTopLngStr("LtxtLoginSecReq") & ""
sql = "select T0.SecID, IsNull(T1.AlterSecName, T0.SecName) SecName, " & _
"T0.SecContent, T0.ReqLogin, T0.Status, T0.Type, " & _
"Case T0.Type When 'L' Then T0.SecContent Else '' End Link, Case T0.Type When 'R' Then T0.SecContent Else '' End rsIndex, " & _
"Case T0.Type When 'R' Then (select Count('') from OLKRSVars where rsIndex = Convert(int,Convert(nvarchar(100),T0.SecContent))) Else 0 End rsVarCount " & _
"from OLKSections T0 " & _
"left outer join OLKSectionsAlterNames T1 on T1.SecType = T0.SecType and T1.SecID = T0.SecID and T1.LanID = " & Session("LanID") & " " & _
"where T0.HideMainMenu = 'N' and T0.UserType = 'A' and T0.Status = 'A' "

If Session("useraccess") = "U" Then
	If myAut.AuthorizedForms <> "" Then
		sql = sql & "and T0.SecID in (" & myAut.AuthorizedForms & ") "
	Else
		sql = sql & "and 1 = 2 "
	End If
End If
				
sql = sql & "order by T0.SecOrder asc"
set rs = conn.execute(sql) %>
var mnuForms = new Menu('MnuForms', true);
<% do while not rs.eof
If rs("ReqLogin") = "N" or rs("ReqLogin") = "Y" and Session("UserName") <> "" Then
	Select Case rs("Type")
		Case "N"
			mnuLink = "goLink(\'sec.asp?SecID=" & rs("SecID") & "\');"
		Case "L"
			mnuLink = "goLinkPop(\'" & Replace(rs("Link"), "'", "\'") & "\');"
		Case "R"
			mnuLink = "goRep(" & rs("rsIndex") & ", " & rs("rsVarCount") & ");"
	End Select
Else
	mnuLink = "alert(\'" & SecErrMsg & "\');"
End If %>
mnuForms.addMenuItem(new MenuItem('<%=Replace(myHTMLEncode(rs("SecName")), "'", "\'")%>&nbsp;&nbsp;', '<%=mnuLink%>', false, null));
<% rs.movenext
loop %>
<% End If %>

var mnuNew = new Menu('New', false);
<% If myApp.EnableOCLG and Session("UserName") <> "" Then %>mnuNew.addMenuItem(new MenuItem('<%=getagentTopLngStr("DtxtActivity")%>&nbsp;&nbsp;', 'goLink(\'addActivity/goNewActivity.asp?AddPath=\');', false, null));<% End If %>
<% If myApp.EnableOITM Then %>mnuNew.addMenuItem(new MenuItem('<%=getagentTopLngStr("DtxtItem")%>&nbsp;&nbsp;', 'goLink(\'addItem/goNewItem.asp?AddPath=\');', false, null));<% End If %>
<% If myApp.EnableOCRD Then %>mnuNew.addMenuItem(new MenuItem('<%=getagentTopLngStr("DtxtBP")%>&nbsp;&nbsp;', 'goLink(\'addCard/goNewCard.asp?AddPath=\');', false, null));<% End If %>
<% If myApp.EnableOOPR and Session("UserName") <> "" Then %>mnuNew.addMenuItem(new MenuItem('<%=getagentTopLngStr("DtxtSO")%>&nbsp;&nbsp;', 'goLink(\'addSO/goNewSO.asp?AddPath=\');', false, null));<% End If %>
<% If Session("UserName") <> "" and comDocsMenu then %>mnuNew.addMenuItem(new MenuItem('<%=getagentTopLngStr("LtxtComDocs")%>&nbsp;&nbsp;', '#', true, 'NewDocs'));<% End If %>
<% If Session("UserName") <> "" and CardType = "C" and myApp.EnableORCT Then %>mnuNew.addMenuItem(new MenuItem('<%=Replace(myHTMLEncode(txtRct), "'", "\'")%>&nbsp;&nbsp;', 'createPayment();', false, null));<% End If %>
mnuNew.addMenuItem(new MenuItem('<%=getagentTopLngStr("DtxtMessage")%>&nbsp;&nbsp;', 'goLink(\'newMessage.asp\');', false, null));

var mnuOpenDocs = new Menu('OpenDocs', true);
<% If myApp.EnableOCLG Then %>mnuOpenDocs.addMenuItem(new MenuItem('<%=getagentTopLngStr("LtxtOpenActivities")%>&nbsp;&nbsp;', 'goLink(\'openedActivities.asp\');', false, null));<% End If %>
<% If myApp.EnableOOPR Then %>mnuOpenDocs.addMenuItem(new MenuItem('<%=getagentTopLngStr("LtxtOpenSO")%>&nbsp;&nbsp;', 'goLink(\'openedSO.asp\');', false, null));<% End If %>
<% If myApp.EnableOITM Then %>mnuOpenDocs.addMenuItem(new MenuItem('<%=getagentTopLngStr("LtxtOpenItms")%>&nbsp;&nbsp;', 'goLink(\'openedItems.asp\');', false, null));<% End If %>
<% If myApp.EnableOCRD Then %>mnuOpenDocs.addMenuItem(new MenuItem('<%=Replace(getagentTopLngStr("LtxtOpenCrd"),"{0}", myHTMLEncode(txtClients))%>&nbsp;&nbsp;', 'goLink(\'openedCards.asp\');', false, null));<% End If %>
<% If comDocsMenu or myApp.EnableORCT Then %>mnuOpenDocs.addMenuItem(new MenuItem('<%=getagentTopLngStr("LtxtOpenDocs")%>&nbsp;&nbsp;', 'goLink(\'openedDocs.asp\');', false, null));<% End If %>

var mnuGenMan = new Menu('GenMan', true);
<% If myAut.HasAuthorization(4) Then
mnuManage = True %>mnuGenMan.addMenuItem(new MenuItem('<%=getagentTopLngStr("LtxtNewsMngmnt")%>&nbsp;&nbsp;', 'goLink(\'newsman.asp\');', false, null));<% End If %>
<% If myAut.HasAuthorization(38) Then
mnuManage = True %>mnuGenMan.addMenuItem(new MenuItem('<%=getagentTopLngStr("LtxtPollMngmnt")%>&nbsp;&nbsp;', 'goLink(\'pollman.asp\');', false, null));<% End If %>
<% If myAut.HasAuthorization(39) Then
mnuManage = True %>mnuGenMan.addMenuItem(new MenuItem('<%=getagentTopLngStr("LtxtExtPollMngmnt")%>&nbsp;&nbsp;', 'goLink(\'extpollman.asp\');', false, null));<% End If %>

var mnuOp = new Menu('Op', false);
<% If myAut.HasBPAccess Then %>mnuOp.addMenuItem(new MenuItem('<%=Replace(getagentTopLngStr("LtxtClientSearch"), "{0}", myHTMLEncode(txtClients))%>&nbsp;&nbsp;', 'goLink(\'searchClient.asp\');', false, null));
<% If myAut.HasAuthorization(24) and IsBPAssigned Then %>mnuOp.addMenuItem(new MenuItem('<%=myHTMLDecode(getagentTopLngStr("LtxtStateOfAcct"))%>&nbsp;&nbsp;', 'goLink(\'cxc.asp\');', false, null));<% End If %><% End If %>
<% If EnableSections Then %>mnuOp.addMenuItem(new MenuItem('<%=getagentTopLngStr("DtxtForms")%>&nbsp;&nbsp;', '#', true, 'MnuForms'));<% End If %>
<% If comDocsMenu or myApp.EnableOCLG or myApp.EnableOITM or myApp.EnableOCRD or myApp.EnableORCT Then %>mnuOp.addMenuItem(new MenuItem('<%=getagentTopLngStr("LtxtOpenObjs")%>&nbsp;&nbsp;', '#', true, 'OpenDocs'));<% End If %>
<% If mnuManage Then %>mnuOp.addMenuItem(new MenuItem('<%=getagentTopLngStr("LtxtGenMan")%>&nbsp;', '#', 'true', 'GenMan'));<% End If %>
mnuOp.addMenuItem(new MenuItem('<%=getagentTopLngStr("LtxtChangePwd")%>', 'goLink(\'changePwd.asp\');', false, null));
<% If Session("UserAccess") = "P" Then %>mnuOp.addMenuItem(new MenuItem('<%=getagentTopLngStr("LtxtRecover")%>', 'goLink(\'recoverSearch.asp\');', false, null));<% End If %>

var mnuOlkConf = new Menu('OlkConf', true);
<% If Session("useraccess") = "P" or Session("HasActionConfAut") Then %>mnuOlkConf.addMenuItem(new MenuItem('<%=getagentTopLngStr("LtxtActions")%>&nbsp;&nbsp;', 'doMyLink(\'executeConf.asp\', \'Type=A\', \'_self\');', false, null));<% End If %>
<% If myAut.HasAuthorization(114) or myAut.HasAuthorization(116) or myAut.HasAuthorization(118) Then %>mnuOlkConf.addMenuItem(new MenuItem('<%=getagentTopLngStr("DtxtBP")%>&nbsp;&nbsp;', 'doMyLink(\'executeConf.asp\', \'Type=C\', \'_self\');', false, null));<% End If %>
<% If myAut.HasAuthorization(121) Then %>mnuOlkConf.addMenuItem(new MenuItem('<%=getagentTopLngStr("DtxtItem")%>&nbsp;&nbsp;', 'doMyLink(\'executeConf.asp\', \'Type=I\', \'_self\');', false, null));<% End If %>
<% If myAut.HasAuthorization(126) Then %>mnuOlkConf.addMenuItem(new MenuItem('<%=Replace(myHTMLEncode(txtRct), "'", "\'")%>&nbsp;&nbsp;', 'doMyLink(\'executeConf.asp\', \'Type=R\', \'_self\');', false, null));<% End If %>
<% If Session("useraccess") = "P" or Session("HasComDocConf") Then %>mnuOlkConf.addMenuItem(new MenuItem('<%=getagentTopLngStr("LtxtComDocs")%>&nbsp;&nbsp;', 'doMyLink(\'executeConf.asp\', \'Type=D\', \'_self\');', false, null));<% End If %>

<%
autMnu = False
autCount_O = 0
autCount_A = 0
autCount_C = 0
autCount_D = 0
autCount_R = 0

sql = "declare @SlpCode int set @SlpCode = " & Session("vendid") & " " & _  
"select Case When Left(T0.ExecAt, 1) = 'O' Then 'O' Else T0.ExecAt End ExecAt, Count('') [Count] " & _  
"from OLKUAFControl T0   " & _  
"left outer join R3_ObsCommon..TLOG T5 on T5.LogNum = T0.ObjectEntry " & _  
"inner join OLKUAFControl2 X0 on X0.ID = T0.ID " & _  
"inner join OLKUAF4 X1 on X1.FlowID = X0.FlowID and X1.AutGrpID = X0.AutGrpID " & _  
"inner join OLKAutGrpSlp X2 on X2.GrpID = X0.AutGrpID and X2.SlpCode = @SlpCode " & _  
"inner join OLKUAF X3 on X3.FlowID = X0.FlowID " & _  
"where T0.Status = 'O' and (T5.Status = 'H' or LEFT(T0.ExecAt, 1) = 'O') and X0.Status = 'W' " & _  
"and  " & _  
"( " & _  
"	(select Status from OLKUAFControl2 where ID = X0.ID and FlowID = X0.FlowID and LineID = X0.LineID-1) is null " & _  
"	or " & _  
"	(select Status from OLKUAFControl2 where ID = X0.ID and FlowID = X0.FlowID and LineID = X0.LineID-1) = 'A' " & _  
") " & _  
"Group By Case When Left(T0.ExecAt, 1) = 'O' Then 'O' Else T0.ExecAt End " 
set rAut = Server.CreateObject("ADODB.RecordSet")
set rAut = conn.execute(sql)
do while not rAut.eof
	Select Case rAut("ExecAt")
		Case "O"
			autCount_O = rAut("Count")
			autMnu = True
		Case "A1"
			autCount_A = rAut("Count")
			autMnu = True
		Case "C1"
			autCount_C = rAut("Count")
			autMnu = True
		Case "D3"
			autCount_D = rAut("Count")
			autMnu = True
		Case "R2"
			autCount_R = rAut("Count")
			autMnu = True
	End Select
rAut.movenext
loop
%>
var mnuOlkConf = new Menu('OlkAut', true);
<% If autCount_O > 0 Then %>mnuOlkConf.addMenuItem(new MenuItem('<%=getagentTopLngStr("LtxtActions")%>&nbsp;&nbsp;', 'doMyLink(\'executeAut.asp\', \'Type=A\', \'_self\');', false, null));<% End If %>
<% If autCount_C > 0 Then %>mnuOlkConf.addMenuItem(new MenuItem('<%=getagentTopLngStr("DtxtBP")%>&nbsp;&nbsp;', 'doMyLink(\'executeAut.asp\', \'Type=C\', \'_self\');', false, null));<% End If %>
<% If autCount_A > 0 Then %>mnuOlkConf.addMenuItem(new MenuItem('<%=getagentTopLngStr("DtxtItem")%>&nbsp;&nbsp;', 'doMyLink(\'executeAut.asp\', \'Type=I\', \'_self\');', false, null));<% End If %>
<% If autCount_R > 0 Then %>mnuOlkConf.addMenuItem(new MenuItem('<%=Replace(myHTMLEncode(txtRct), "'", "\'")%>&nbsp;&nbsp;', 'doMyLink(\'executeAut.asp\', \'Type=R\', \'_self\');', false, null));<% End If %>
<% If autCount_D > 0 Then %>mnuOlkConf.addMenuItem(new MenuItem('<%=getagentTopLngStr("LtxtComDocs")%>&nbsp;&nbsp;', 'doMyLink(\'executeAut.asp\', \'Type=D\', \'_self\');', false, null));<% End If %>
<% set rAut = Nothing %>
var mnuMyTask = new Menu('MyTask', false);
<% If myApp.EnableOCLG Then %>mnuMyTask.addMenuItem(new MenuItem('<%=getagentTopLngStr("LtxtOpenActivities")%>&nbsp;&nbsp;', 'doMyLink(\'searchOpenedActivities.asp\', \'SlpCodeFrom=<%=AgentName%>&SlpCodeTo=<%=AgentName%>\', \'_self\');', false, null));<% End If %>
<% If myApp.EnableOOPR Then %>mnuMyTask.addMenuItem(new MenuItem('<%=getagentTopLngStr("LtxtOpenSO")%>&nbsp;&nbsp;', 'doMyLink(\'searchOpenedSO.asp\', \'SlpCodeFrom=<%=AgentName%>&SlpCodeTo=<%=AgentName%>\', \'_self\');', false, null));<% End If %>
<% If comDocsMenu or myApp.EnableORCT Then %>mnuMyTask.addMenuItem(new MenuItem('<%=getagentTopLngStr("LtxtOpenDocs")%>&nbsp;&nbsp;', 'doMyLink(\'searchOpenedDocs.asp\', \'orden1=0&orden2=A&SlpCodeFrom=<%=AgentName%>&SlpCodeTo=<%=AgentName%>\', \'_self\');', false, null));<% End If %>
<% If autMnu Then %>mnuMyTask.addMenuItem(new MenuItem('<%=getagentTopLngStr("DtxtAuthorize")%>&nbsp;&nbsp;', '#', 'true', 'OlkAut'));<% End If %>
<% If myAut.HasObjConfirmAccess Then %>mnuMyTask.addMenuItem(new MenuItem('<%=getagentTopLngStr("DtxtConfirm")%>&nbsp;&nbsp;', '#', 'true', 'OlkConf'));<% End If %>
mnuMyTask.addMenuItem(new MenuItem('<%=getagentTopLngStr("LtxtExtendedPolls")%>&nbsp;&nbsp;', 'goLink(\'extPollList.asp\');', false, null));
<% If optOfert and myAut.HasAuthorization(7) Then %>mnuMyTask.addMenuItem(new MenuItem('<%=Replace(getagentTopLngStr("LtxtOfertMan"), "{0}", myHTMLEncode(txtOferts))%>&nbsp;&nbsp;', 'doMyLink(\'ofertsMan.asp\', \'dtBy=O&OfertStatus=W, O&orden1=8&orden2=desc&SlpCodeFrom=<%=AgentName%>&SlpCodeTo=<%=AgentName%>\', \'_self\');', false, null));<% End If %>
<% If EnableClientActivation and myAut.HasAuthorization(89) Then %>mnuMyTask.addMenuItem(new MenuItem('<%=getagentTopLngStr("LtxtAnonRegActive")%>&nbsp;&nbsp;', 'goLink(\'activation.asp\');', false, null));<% End If %>
mnuMyTask.addMenuItem(new MenuItem('<%=getagentTopLngStr("LtxtMsgBox")%>&nbsp;&nbsp;', 'goLink(\'agent.asp?onlyMsg=Y\');', false, null));

var mnuRep = new Menu('Rep', false);
<% 
set rsRepRG = Server.CreateObject("ADODB.RecordSet")

If Session("useraccess") = "U" Then
	If myAut.AuthorizedRepGroups <> "" Then
		sqlAdd = "and T0.rgIndex in (" & myAut.AuthorizedRepGroups & ") "
	Else
		sqlAdd = " and 1 = 2 "
	End If
Else
	sqlAdd = ""
End If

sql = "select Case T0.rgIndex When -2 Then 'A' Else 'B' End Ordr2, T0.rgIndex, IsNull(alterRGName, rgName) rgName " & _
"from olkRG T0 " & _
"left outer join OLKRGAlterNames T1 on T1.rgIndex = T0.rgIndex and T1.LanID = " & Session("LanID") & " " & _
"where UserType = 'V' and exists (select 'A' from OLKRS where rgIndex = T0.rgIndex and Active = 'Y' and LinkOnly = 'N') " & sqlAdd & " order by Ordr2, rgName asc"
rsRepRG.open sql, conn, 3, 1
If not rsRepRG.eof Then
do while not rsRepRG.eof %>
mnuRep.addMenuItem(new MenuItem('<%=Replace(myHTMLEncode(rsRepRG("rgName")), "'", "\'")%>&nbsp;&nbsp;', '#', true, 'RG<%=rsRepRG.bookmark%>'));
<% rsRepRG.movenext
loop
rsRepRG.movefirst
End If %>


<% 
If Not rsRepRG.eof Then
set rsRep = Server.CreateObject("ADODB.RecordSet")
do while not rsRepRG.eof
sql = "select T0.rsIndex, IsNull(alterRSName, rsName) rsName, " & _
"(select count('A') from olkRSVars where rsIndex = T0.rsIndex)+Case rsTop When 'Y' Then 1 Else 0 End varsCount " & _
"from olkRS T0 " & _
"left outer join OLKRSAlterNames T1 on T1.rsIndex = T0.rsIndex and T1.LanID = " & Session("LanID") & " " & _
"where T0.rgIndex = " & rsRepRG("rgIndex") & " and Active = 'Y' and LinkOnly = 'N' order by rsName asc "
rsRep.open sql, conn, 3, 1
%>
var mnuRG<%=rsRepRG.bookmark%> = new Menu('RG<%=rsRepRG.bookmark%>', true);
<% do while not rsRep.eof %>
mnuRG<%=rsRepRG.bookmark%>.addMenuItem(new MenuItem('<%=Replace(myHTMLEncode(rsRep("rsName")), "'", "\'")%>&nbsp;&nbsp;', 'goRep(<%=rsRep("rsIndex")%>, <%=rsRep("varsCount")%>);', false, null));
<% rsRep.movenext
loop %>
<% rsRep.close
rsRepRG.movenext
loop 
rsRepRG.movefirst
set rsRep = nothing
End If %>
	
var mnuHelp = new Menu('Help', false);
mnuHelp.addMenuItem(new MenuItem('<%=getagentTopLngStr("LtxtAbout")%>...&nbsp;&nbsp;', 'OpenWin = window.open(\'help/about.asp?Version=<%=rh("Version")%>\', \'AboutUs\', \'toolbar=no,menubar=no,location=no,scrollbars=no,resizable=no, width=520,height=347\');', false, null));	

</script>

<script language="javascript">loadMenus();</script>

<STYLE TYPE="TEXT/CSS">
<!--
body{
scrollbar-base-color:#014891;
scrollbar-face-color:#0069D2;
scrollbar-highlight-color:#0069D2;
scrollbar-3dlight-color:#014891;
scrollbar-darkshadow-color:#014891;
scrollbar-Shadow-color:#014891;
scrollbar-arrow-color:#FFFFFF;
scrollbar-track-color:#0068D1;
}
.input		
{

	
	color : #3366CC;
	font-family : Verdana, Arial, Helvetica, sans-serif;
	font-size : 10px;
	background-image: url('menybg.gif');
	background-repeat: repeat-x;
	border: 1px solid #555555
}
-->
</STYLE>
<script language="javascript">
var OpenWin;
var curDir = '<% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>';
function loadStartUp()
{
<% If Request("err") = "inv" Then %>
	alert("<%=getagentTopLngStr("LtxtErrItmInv")%>")
<% ElseIf Request("err") = "tax" Then
	RequestVal = ""
		For each itm in Request.Form
			If itm <> "err" and itm <> "tItem" and itm <> "cmd" and itm <> "Item" and itm <> "DocFlowErr" Then
				RequestVal = RequestVal & "{y}" & itm & "{i}" & Request(itm)
			End If
		Next
		For each itm in Request.QueryString
			If itm <> "err" and itm <> "tItem" and itm <> "cmd" and itm <> "Item" and itm <> "DocFlowErr" Then
				RequestVal = RequestVal & "{y}" & itm & "{i}" & Request(itm)
			End If
		Next %>
	OpenWin = this.open("", "GetTaxCode", "toolbar=no,menubar=no,location=no,scrollbars=no,resizable=no,width=300,height=20,top="+wint+",left="+winl);
	doMyLink('cart/AddCartGetTaxCode.asp', 'Item=<%=Request("tItem")%>&expItem=<%=Request("expItem")%>&AddPath=&pop=Y&redir=<%=curCmd%>&retVal=<%=RequestVal%>', 'GetTaxCode');
<% ElseIf Request("errMInv") <> "" Then %>
	var alertMsg = '<%=getagentTopLngStr("LtxtErrMultItmInv")%>: \n<%=Replace(Request("errMInv"), "'", "\'")%>'
	alert(alertMsg);
<% end if %>
<% If Request("DocFlowErr") <> "" Then %>
	OpenWin = this.open('', "CtrlWindow", "toolbar=no,menubar=no,location=no,scrollbars=yes,resizable=no,width=550,height=400,top="+wint+",left="+winl);
	doMyLink('flowAlert.asp', 'DocFlowErr=<%=Request("DocFlowErr")%>&pop=Y&cmd=<%=curCmd%>&Item=<%=Request("Item")%>&addQty=<%=Request("addQty")%>&precio=<%=Request("precio")%>&retURL=<%=Request("retURL")%>', 'CtrlWindow');
<% End If %>
<% Select Case Request("LicErr")
	Case "TMNOOP"    ' Licencia Rechazada %>
	alert('<%=getagentTopLngStr("LtxtLicRejected")%>');
<%	Case "TMNOLIC" ' No cuenta con licencia %>
	alert('<%=getagentTopLngStr("LtxtNoLic")%>');
<%	Case "TMPSNLIC" ' Licencias usadas %>
	alert('<%=getagentTopLngStr("LtxtLicLimit")%>');
<% 	Case "TMLICVENCI"   ' Licencia vencida %>
	alert('<%=getagentTopLngStr("LtxtLicExp")%>');
<%	Case "NO" %>
	alert('<%=getagentTopLngStr("LtxtLicReadOnly")%>');
<% End Select %>
<% If Request("fastAddErr") = "Y" Then %>
<% Select Case Request("fastAddErrType")
	Case "D" %>
	alert('<%=getagentTopLngStr("LtxtFastAddErrDisc")%>');
<% Case "F" %>
	alert('<%=getagentTopLngStr("LtxtFastAddErrBlock")%>'.replace('{0}', '<%=Request("fastAddErrItm")%>'));
<%	Case Else %>
	alert('<%=getagentTopLngStr("LtxtFastAddErr")%>'.replace('{0}', '<%=Request("fastAddErrItm")%>'));
	<% End Select %>
<% ElseIf Request("loadRec") <> "" Then
	lineID = Request("ItmEntry") & Request("ItmEntry") %>
	ItemCmd = '<% If Session("RetVal") = "" Then %>D<% Else %>A<% End If %>';
	openItemDetails('<%=Replace(Request("loadRec"), "'", "\'")%>');	
<% End If %>
}
</script>
<link rel="stylesheet" type="text/css" href="design/0/style/stylenuevo.css">
<link rel="stylesheet" type="text/css" media="all" href="design/0/style/style_cal.css" title="winter" />
<% SelDes = "0" %>
</head>
<%
OnLoadFocus = "loadStartUp();startBlink();"
hasFocus = False
Select Case strScriptName
	Case "search.asp" 
		hasFocus = True
		OnLoadFocus = OnLoadFocus & "document.frmSmallSearch.string.focus();" 
	Case "searchclient.asp"
		hasFocus = True
		OnLoadFocus = OnLoadFocus & "document.frmSearchOCRD.string.focus();"
	Case "agentsearch.asp"
		hasFocus = True
		OnLoadFocus = OnLoadFocus & "document.frmSmallSearch.string.focus();"
End Select

If Request("focus") <> "" Then
	hasFocus = True
	OnLoadFocus = OnLoadFocus & "document." & Request("focus") & ".focus();"
End If

If Not hasFocus and Session("RetVal") <> "" and (strScriptName = "cart.asp" or strScriptName = "search") Then
	OnLoadFocus = OnLoadFocus & "document.frmSmallSearch.string.focus();" 
End If

If Session("RetVal") <> "" and (strScriptName = "cart.asp" or strScriptName = "search.asp" or strScriptName = "adcustomsearch.asp" or strScriptName = "wish.asp") Then OnLoadFocus = OnLoadFocus & "setMinRepSize();"
%>
<body topmargin="0" leftmargin="0" link="#4783C5" vlink="#4783C5" onfocus="chkWin()" bgcolor="#0166CB" onload="javascript:chkWin();<% If strScriptName = "report.asp" Then %>setInterval('doBlink();',1000);<% End If %><%=OnLoadFocus%>" onresize="javascript:setAlertPos();" style="background-color: #0066CA">
<!--#include file="licid.inc"-->
<table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
  <tr>
    <td>
    <img src="ventas/images/spacer.gif" width="92" height="1" border="0" alt></td>
    <td>
    <img src="ventas/images/spacer.gif" width="24" height="1" border="0" alt></td>
    <td>
    <img src="ventas/images/spacer.gif" width="47" height="1" border="0" alt></td>
    <td><img src="ventas/images/spacer.gif" border="0" alt></td>
    <td>
    <img src="ventas/images/spacer.gif" width="22" height="1" border="0" alt></td>
    <td>
    <img src="ventas/images/spacer.gif" width="11" height="1" border="0" alt></td>
    <td>
    <img src="ventas/images/spacer.gif" width="56" height="1" border="0" alt></td>
    <td><img src="ventas/images/spacer.gif" border="0" alt></td>
    <td>
    <img src="ventas/images/spacer.gif" width="150" height="1" border="0" alt></td>
    <td><img src="ventas/images/spacer.gif" border="0" alt></td>
    <td>
    <img src="ventas/images/spacer.gif" width="73" height="1" border="0" alt></td>
    <td><img src="ventas/images/spacer.gif" border="0" alt></td>
    <td>
    <img src="ventas/images/spacer.gif" width="88" height="1" border="0" alt></td>
  </tr>
  <tr id="pageTop">
    <td colspan="5">
    <a href="agent.asp">
    <img name="art_ventas1_r1_c1" src="ventas/images/<%=Session("rtl")%>art_ventas1_r1_c1.jpg" border="0" alt longdesc="<%=getagentTopLngStr("LtxtBackToPortal")%>"></a></td>
    <td colspan="8" background="ventas/images/art_ventas1_r1_c14.jpg" width="100%" valign="top">
    <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
      <tr>
        <td height="90" valign="middle" align="center" onclick="javascript:window.location.href='agent.asp';" style="cursor: hand; background-image: url('ventas/images/<%=Session("rtl")%>art_ventas1_r1_c6.jpg'); background-repeat: no-repeat; padding-<% If Session("rtl") = "" Then %>left<% Else %>right<% End If %>: 190px; <% If Session("rtl") <> "" Then %>background-position: top right;<% End If %>">
    	<% If Not IsNull(myApp.AgentLogo) Then %><img src="imagenes/<%=Session("olkdb")%>/<%=myApp.AgentLogo%>"><% Else %>&nbsp;<% End If %></td>
        <td width="300" valign="bottom">
        <table border="0" cellpadding="0" cellspacing="0"  bordercolor="#111111" width="100%" style="border-collapse: collapse">
          <tr>
            <td width="100%">
            <div align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
              <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111">
                <tr>
                  <td>
                  <a href="login.asp?logout=Y">
                  <img border="0" src="ventas/images/lockmini.gif"></a></td>
                  <td>
					<b><font size="1" face="Verdana">
					<a href="login.asp?logout=Y<% If myApp.EnableAnSesion Then %>&amp;redir=<% End If %>"><font color="#8DC6FE" face="Verdana" size="1">
					<%=getagentTopLngStr("LtxtLogOut")%></font></a></font></b></td>
                </tr>
              </table>
            </div>
            </td>
          </tr>
          <tr>
            <td width="100%" align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b>
            <font size="1" face="Verdana" color="#8DC6FE"><%=doFormatDate(Date)%>
			&nbsp;</font></b></td>
          </tr>
          <tr>
            <td width="100%" align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
            <font size="1" face="Verdana" color="#8DC6FE"><% If Session("branch") <> -1 Then %>
			(<%=rh("branchName")%>)&nbsp;<%end if%><b><%=mySession.GetCompanyName%></b><br><b><% If 1 = 2 Then %>
			Agente<% Else %><%=txtAgent%><% End If %>: <%=AgentName%>&nbsp;</b></font></td>
          </tr>
          <tr>
            <td width="100%" align="right" height="5"></td>
          </tr>
          </table>
        </td>
      </tr>
    </table>
    </td>
  </tr>
  <tr>
    <td colspan="13" background="ventas/images/riega_arriba.jpg">
<table border="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
  <tr>
    <td width="20" background="ventas/images/art_ventas1_r2_c1.jpg">&nbsp;
    </td>
    <td width="1">
    <font face="Verdana">
    <img name="art_ventas1_r2_c4" src="ventas/images/art_ventas1_r2_c4.jpg" border="0" alt></font></td>
    <td class="myMnuItem" onmouseover="javascript:showMnu(this, 'New', false);" onmouseout="javascript:hideMnu();" width="80">
    <p align="center">
    <b><font color="#FFFFFF" size="1" face="Verdana"><%=getagentTopLngStr("DtxtNew")%></font></b></td>
    <td width="1">
    <font face="Verdana">
    <img name="art_ventas1_r2_c4" src="ventas/images/art_ventas1_r2_c4.jpg" border="0" alt></font></td>
    <td class="myMnuItem" onmouseover="javascript:showMnu(this, 'MyTask', false);" onmouseout="javascript:hideMnu();" width="80">
    <p align="center"><b><font color="#FFFFFF" size="1" face="Verdana">
	<%=getagentTopLngStr("LtxtMyTasks")%></font></b></td>
    <td width="1">
    <font face="Verdana">
    <img name="art_ventas1_r2_c4" src="ventas/images/art_ventas1_r2_c4.jpg" border="0" alt></font></td>
    <td class="myMnuItem" onmouseover="javascript:showMnu(this, 'Op', false);" onmouseout="javascript:hideMnu();" width="80">
    <p align="center"><b><font color="#FFFFFF" size="1" face="Verdana">
	<%=getagentTopLngStr("LtxtOperations")%></font></b></td>
    <% If myAut.HasAuthorization(1) Then %>
    <td width="1">
    <font face="Verdana">
    <img name="art_ventas1_r2_c8" src="ventas/images/art_ventas1_r2_c8.jpg" border="0" alt></font></td>
    <td class="myMnuItem" onmouseover="javascript:showMnu(this, 'Cat', false);" onmouseout="javascript:hideMnu();" style="cursor: hand" onclick="javascript:window.location.href='agentSearch.asp';" width="80">
    <p align="center"><b><font color="#FFFFFF" face="Verdana" size="1">
    <span style="text-decoration: none"><%=getagentTopLngStr("DtxtCat")%></span></font></b></td><% End If %>
    <td width="1">
    <font face="Verdana">
    <img name="art_ventas1_r2_c10" src="ventas/images/art_ventas1_r2_c10.jpg" border="0" alt></font></td>
    <td class="myMnuItem" onmouseover="javascript:showMnu(this, 'Rep', false);" onmouseout="javascript:hideMnu();" width="80">
    <p align="center"><b><font face="Verdana" size="1" color="#FFFFFF">
	<%=getagentTopLngStr("LtxtReps")%></font></b></td>
    <td width="1">
    <font face="Verdana">
    <img name="art_ventas1_r2_c10" src="ventas/images/art_ventas1_r2_c10.jpg" border="0" alt></font></td>
    <td class="myMnuItem" onmouseover="javascript:showMnu(this, 'Help', false);" onmouseout="javascript:hideMnu();" width="80">
    <p align="center"><b><font color="#FFFFFF" size="1" face="Verdana">
	<%=getagentTopLngStr("LtxtHelp")%></font></b></td>
    <td width="1">
    <img name="art_ventas1_r2_c12" src="ventas/images/art_ventas1_r2_c12.jpg" border="0" alt></td>
    <td background="ventas/images/art_ventas1_r2_c13.jpg" valign="baseline" align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
    <table border="0" cellpadding="0" cellspacing="0">
    <% If Session("UserName") <> "" Then %><td><font size="1" face="Verdana" color="#8DC6FE"><b>
		<%=getagentTopLngStr("LtxtCurClient")%>: <%=rh("CardName")%>&nbsp;</b></font></td><% End If %>
	<td><% If 1 = 2 Then %><a href="#" onclick="javascript:OpenWin = window.open('help/help.asp?cmd=<%=curCmd%>', 'olkHELP', 'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,width=400, height=600' )">
	<img border="0" src="ventas/images/help2.gif" width="18" height="17"></a><% Else %>&nbsp;<% End If %></td></table></td>
  </tr>
</table>
    </td>
  </tr>
  <tr>
    <td colspan="5">
    <img name="art_conimg_r3_c1" src="ventas/images/<%=Session("rtl")%>art_conimg_r3_c1.jpg" border="0" alt></td>
    <td colspan="8" background="ventas/images/art_ventas1_r3_c14.jpg" valign="top">
    <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
      <tr>
        <td width="50%">
    <table border="0" cellpadding="0" cellspacing="0" width="381">
      <!-- fwtable fwsrc="\\Sboserver\Userh\pablomorano\topmanage\logos\originales\iconos_ventas.png" fwbase="iconos_ventas.gif" fwstyle="FrontPage" fwdocid = "742308039" fwnested=""0" -->
      <script language="JavaScript">
	  <!-- hide 
	  if (document.images) {
	  iconos_ventas_r2_c2_f2 = new Image(31 ,27); iconos_ventas_r2_c2_f2.src = "iconos_ventas_r2_c2_f2.gif";
	  iconos_ventas_r2_c2_f1 = new Image(31 ,27); iconos_ventas_r2_c2_f1.src = "iconos_ventas_r2_c2.gif";
	  iconos_ventas_r2_c4_f2 = new Image(51 ,27); iconos_ventas_r2_c4_f2.src = "iconos_ventas_r2_c4_f2.gif";
	  iconos_ventas_r2_c4_f1 = new Image(51 ,27); iconos_ventas_r2_c4_f1.src = "iconos_ventas_r2_c4.gif";
	  iconos_ventas_r2_c6_f2 = new Image(31 ,27); iconos_ventas_r2_c6_f2.src = "iconos_ventas_r2_c6_f2.gif";
	  iconos_ventas_r2_c6_f1 = new Image(31 ,27); iconos_ventas_r2_c6_f1.src = "iconos_ventas_r2_c6.gif";
	  iconos_ventas_r2_c8_f2 = new Image(31 ,27); iconos_ventas_r2_c8_f2.src = "iconos_ventas_r2_c8_f2.gif";
	  iconos_ventas_r2_c8_f1 = new Image(31 ,27); iconos_ventas_r2_c8_f1.src = "iconos_ventas_r2_c8.gif";
	  iconos_ventas_r2_c10_f2 = new Image(31 ,27); iconos_ventas_r2_c10_f2.src = "iconos_ventas_r2_c10_f2.gif";
	  iconos_ventas_r2_c10_f1 = new Image(31 ,27); iconos_ventas_r2_c10_f1.src = "iconos_ventas_r2_c10.gif";
	  iconos_ventas_r2_c12_f2 = new Image(31 ,27); iconos_ventas_r2_c12_f2.src = "iconos_ventas_r2_c12_f2.gif";
	  iconos_ventas_r2_c12_f1 = new Image(31 ,27); iconos_ventas_r2_c12_f1.src = "iconos_ventas_r2_c12.gif";
	  iconos_ventas_r2_c14_f2 = new Image(31 ,27); iconos_ventas_r2_c14_f2.src = "iconos_ventas_r2_c14_f2.gif";
	  iconos_ventas_r2_c14_f1 = new Image(31 ,27); iconos_ventas_r2_c14_f1.src = "iconos_ventas_r2_c14.gif";
	  iconos_ventas_r2_c16_f2 = new Image(31 ,27); iconos_ventas_r2_c16_f2.src = "iconos_ventas_r2_c16_f2.gif";
	  iconos_ventas_r2_c16_f1 = new Image(31 ,27); iconos_ventas_r2_c16_f1.src = "iconos_ventas_r2_c16.gif";
	  }
	  // stop hiding -->
      </script>
      <tr>
        <td>
        <img src="ventas/images/spacer.gif" width="14" height="1" border="0" alt></td>
        <td>
        <img src="ventas/images/spacer.gif" width="31" height="1" border="0" alt></td>
        <td>
        <img src="ventas/images/spacer.gif" width="2" height="1" border="0" alt></td>
        <td>
        <img src="ventas/images/spacer.gif" width="51" height="1" border="0" alt></td>
        <td>
        <img src="ventas/images/spacer.gif" width="2" height="1" border="0" alt></td>
        <td>
        <img src="ventas/images/spacer.gif" width="31" height="1" border="0" alt></td>
        <td>
        <img src="ventas/images/spacer.gif" width="2" height="1" border="0" alt></td>
        <td>
        <img src="ventas/images/spacer.gif" width="31" height="1" border="0" alt></td>
        <td>
        <img src="ventas/images/spacer.gif" width="2" height="1" border="0" alt></td>
        <td>
        <img src="ventas/images/spacer.gif" width="31" height="1" border="0" alt></td>
        <td>
        <img src="ventas/images/spacer.gif" width="2" height="1" border="0" alt></td>
        <td>
        <img src="ventas/images/spacer.gif" width="31" height="1" border="0" alt></td>
        <td>
        <img src="ventas/images/spacer.gif" width="2" height="1" border="0" alt></td>
        <td>
        <img src="ventas/images/spacer.gif" width="31" height="1" border="0" alt></td>
        <td>
        <img src="ventas/images/spacer.gif" width="2" height="1" border="0" alt></td>
        <td>
        <img src="ventas/images/spacer.gif" width="31" height="1" border="0" alt></td>
        <td>
        <img src="ventas/images/spacer.gif" width="85" height="1" border="0" alt></td>
      </tr>
      <tr>
        <td colspan="17">
        <img name="iconos_ventas_r1_c1" src="ventas/images/iconos_ventas_r1_c1.gif" width="381" height="1" border="0" alt></td>
      </tr>
      <tr>
        <td rowspan="2">
        <img name="iconos_ventas_r2_c1" src="ventas/images/iconos_ventas_r2_c1.gif" width="14" height="45" border="0" alt></td>
        <td><% If Not excell and Not pdfDoc Then %><img name="iconos_ventas_r2_c2" src="ventas/images/iconos_ventas_export_disabled.gif" width="31" height="27" border="0" alt="<%=getagentTopLngStr("DtxtExport")%>"><% Else %><a onMouseOut="expTblID=window.setTimeout('clearSmallExpTbl()', 1000);" onMouseOver="MM_swapImage('iconos_ventas_r2_c2','','ventas/images/iconos_ventas_export_over.gif',1);expTblID=window.clearInterval(expTblID);showSmallExpTbl();" href="#"><img name="iconos_ventas_r2_c2" id="imgVentasExport" src="ventas/images/iconos_ventas_export.gif" width="31" height="27" border="0" alt="<%=getagentTopLngStr("DtxtExport")%>"></a><% End If %></td>
        <td rowspan="2">
        <img name="iconos_ventas_r2_c3" src="ventas/images/iconos_ventas_r2_c3.gif" width="2" height="45" border="0" alt></td>
        <td><% If Session("RetVal") <> "" and strScriptName <> "cartsubmit.asp" and strScriptName <> "cartcancel.asp" and myAut.HasAuthorization(26) Then%><a href="javascript:goWL(document, '<%=myApp.GetDefCatOrdr%>', '');" onMouseOut="MM_swapImgRestore();" onMouseOver="MM_swapImage('iconos_ventas_r2_c4','','ventas/images/iconos_ventas_r2_c4_f2.gif',1);"><img name="iconos_ventas_r2_c4" src="ventas/images/iconos_ventas_r2_c4.gif" width="51" height="27" border="0" alt="<%=getagentTopLngStr("LtxtWishList")%>"></a><% else %><img name="iconos_ventas_r2_c4" src="ventas/images/menuicon_arribaoffline_r2_c4.gif" width="51" height="27" border="0" alt="<%=getagentTopLngStr("LtxtWishList")%>"><% end if %></td>
        <td rowspan="2">
        <img name="iconos_ventas_r2_c5" src="ventas/images/iconos_ventas_r2_c5.gif" width="2" height="45" border="0" alt></td>
        <td><% If Session("UserName") <> "" and myAut.HasAuthorization(24) and IsBPAssigned and ClientShowBalance Then %><a href="cxc.asp" onMouseOut="MM_swapImgRestore();" onMouseOver="MM_swapImage('iconos_ventas_r2_c6','','ventas/images/iconos_ventas_r2_c6_f2.gif',1);">
        <p>
        <img name="iconos_ventas_r2_c6" src="ventas/images/iconos_ventas_r2_c6.gif" width="31" height="27" border="0" alt="<%=myHTMLDecode(getagentTopLngStr("LtxtStateOfAcct"))%>"></a><% else %><img name="iconos_ventas_r2_c6" src="ventas/images/menuicon_arribaoffline_r2_c6.gif" width="31" height="27" border="0" alt="<%=myHTMLDecode(getagentTopLngStr("LtxtStateOfAcct"))%>"><% end if %></td>
        <td rowspan="2">
        <img name="iconos_ventas_r2_c7" src="ventas/images/iconos_ventas_r2_c7.gif" width="2" height="45" border="0" alt></td>
        <td><% If EnablePrint Then %><a href="#" onMouseOut="MM_swapImgRestore();" onMouseOver="MM_swapImage('iconos_ventas_r2_c8','','ventas/images/iconos_ventas_r2_c8_f2.gif',1);" onClick="javascript:<% If strScriptName <> "search.asp" then %>printStory('printThis', <%=SelDes%>); return false;<% Else %>printCat('N')<% End If %>"><img name="iconos_ventas_r2_c8" src="ventas/images/iconos_ventas_r2_c8.gif" width="31" height="27" border="0" alt="<%=getagentTopLngStr("LtxtPrint")%>"></a><% Else %><img src="ventas/images/menuicon_arribaoffline_r2_c8.gif" border="0" alt="<%=getagentTopLngStr("LtxtPrint")%>"><% End If %></td>
        <td rowspan="2">
        <img name="iconos_ventas_r2_c9" src="ventas/images/iconos_ventas_r2_c9.gif" width="2" height="45" border="0" alt></td>
        <td><% If Session("UserName") <> "" and optOfert then %><a href="activeClient.asp?open=oferts" onMouseOut="MM_swapImgRestore();" onMouseOver="MM_swapImage('iconos_ventas_r2_c10','','ventas/images/iconos_ventas_r2_c10_f2.gif',1);"><img alt="<%=myHTMLEncode(txtOferts)%>" name="iconos_ventas_r2_c10" src="ventas/images/iconos_ventas_r2_c10.gif" width="31" height="27" border="0" alt></a><%else%><img name="iconos_ventas_r2_c10" src="ventas/images/menuicon_arribaoffline_r2_c10.gif" width="31" height="27" border="0" alt="<%=myHTMLEncode(txtOferts)%>"><%end if%></td>
        <td rowspan="2">
        <img name="iconos_ventas_r2_c11" src="ventas/images/iconos_ventas_r2_c11.gif" width="2" height="45" border="0" alt></td>
        <td><% If session("RetVal") <> "" and strScriptName <> "submitcart.asp" and strScriptName <> "cartsubmit.asp" and strScriptName <> "cartcancel.asp" Then %><a href="cart.asp" onMouseOut="MM_swapImgRestore();" onMouseOver="MM_swapImage('iconos_ventas_r2_c12','','ventas/images/iconos_ventas_r2_c12_f2.gif',1);"><img name="iconos_ventas_r2_c12" src="ventas/images/iconos_ventas_r2_c12.gif" width="31" height="27" border="0" alt="<%=getagentTopLngStr("LtxtCart")%>"></a><%else%><img name="iconos_ventas_r2_c12" src="ventas/images/menuicon_arribaoffline_r2_c12.gif" width="31" height="27" border="0" alt="<%=getagentTopLngStr("LtxtCart")%>"><% end if %></td>
        <td rowspan="2">
        <img name="iconos_ventas_r2_c13" src="ventas/images/iconos_ventas_r2_c13.gif" width="2" height="45" border="0" alt></td>
        <td>
        <% If myAut.HasBPAccess Then %><a href="searchClient.asp" onMouseOut="MM_swapImgRestore();" onMouseOver="MM_swapImage('iconos_ventas_r2_c14','','ventas/images/iconos_ventas_r2_c14_f2.gif',1);">
        <img name="iconos_ventas_r2_c14" src="ventas/images/iconos_ventas_r2_c14.gif" width="31" height="27" border="0" alt="<%=getagentTopLngStr("LtxtClientSearch")%>"></a><% Else %><img name="iconos_ventas_r2_c14" src="ventas/images/menuicon_arribaoffline_r2_c14.gif" width="31" height="27" border="0" alt="<%=getagentTopLngStr("LtxtClientSearch")%>"><% End If %></td>
        <td rowspan="2">
        <img name="iconos_ventas_r2_c15" src="ventas/images/iconos_ventas_r2_c15.gif" width="2" height="45" border="0" alt></td>
        <td><% If Session("Username") <> "" Then %><a onMouseOut="MM_swapImgRestore();" onMouseOver="MM_swapImage('iconos_ventas_r2_c16','','ventas/images/iconos_ventas_r2_c16_f2.gif',1);" href="javascript:goOp('<%=Replace(myHTMLEncode(Session("UserName")), "'", "\'")%>');"><img name="iconos_ventas_r2_c16" src="ventas/images/iconos_ventas_r2_c16.gif" width="31" height="27" border="0" alt="<%=getagentTopLngStr("LtxtOpWithClient")%>"></a><%else%><img name="iconos_ventas_r2_c12" src="ventas/images/menuicon_arribaoffline_r2_c16.gif" width="31" height="27" border="0" alt="<%=getagentTopLngStr("LtxtOpWithClient")%>"><%end if%></td>
        <td rowspan="2">
        <img name="iconos_ventas_r2_c17" src="ventas/images/<%=Session("rtl")%>iconos_ventas_r2_c17.gif" width="85" height="45" border="0" alt></td>
      </tr>
      <tr>
        <td>
        <img name="iconos_ventas_r3_c2" src="ventas/images/iconos_ventas_r3_c2.gif" width="31" height="18" border="0" alt></td>
        <td>
        <img name="iconos_ventas_r3_c4" src="ventas/images/iconos_ventas_r3_c4.gif" width="51" height="18" border="0" alt></td>
        <td>
        <img name="iconos_ventas_r3_c6" src="ventas/images/iconos_ventas_r3_c6.gif" width="31" height="18" border="0" alt></td>
        <td>
        <img name="iconos_ventas_r3_c8" src="ventas/images/iconos_ventas_r3_c8.gif" width="31" height="18" border="0" alt></td>
        <td>
        <img name="iconos_ventas_r3_c10" src="ventas/images/iconos_ventas_r3_c10.gif" width="31" height="18" border="0" alt></td>
        <td>
        <img name="iconos_ventas_r3_c12" src="ventas/images/iconos_ventas_r3_c12.gif" width="31" height="18" border="0" alt></td>
        <td>
        <img name="iconos_ventas_r3_c14" src="ventas/images/iconos_ventas_r3_c14.gif" width="31" height="18" border="0" alt></td>
        <td>
        <img name="iconos_ventas_r3_c16" src="ventas/images/iconos_ventas_r3_c16.gif" width="31" height="18" border="0" alt></td>
      </tr>
    </table>
        </td>
        <td align="<% If Session("rtl") <> "rtl/" Then %>right<% Else %>left<% End If %>">
        	<table cellpadding="0" cellspacing="2" border="0" width="100%">
				<tr>
					<% If strScriptName = "search.asp" Then %>
					<td align="<% If Session("rtl") <> "rtl/" Then %>right<% Else %>left<% End If %>" style="padding-top: 6px;">
					<%
						If Request("document") = "" Then document = ItemDefView Else document = Request("document")
						set objViewType = New clsViewType
						objViewType.ID = "document"
						objViewType.Value = document
						objViewType.OnClick = "document.frmGPage.document.value='{Type}';goPage(1);"
						objViewType.HandCursor = True
						objViewType.doViewType
					%>
					</td>
					<% End If %>
					<td align="<% If Session("rtl") <> "rtl/" Then %>right<% Else %>left<% End If %>" style="padding-top: 4px;<% If strScriptName = "search.asp" Then %>width: 30px;<% End If %>">
					<table cellpadding="0" cellspacing="0" border="0" bgcolor="#0071E3" style="cursor: hand" onclick="return !showSelectLang(this, event);">
						<tr>
							<td width="16" height="16" align="center">
							<font size="1" face="Verdana" color="#FFFFFF"><%=LEFT(UCase(Session("myLng")), 2)%></font></td>
						</tr>
					</table>
					</td>
				</tr>
				<% If strScriptName = "cart.asp" or strScriptName = "agentclient.asp" Then %>
				<tr>
					<td align="<% If Session("rtl") <> "rtl/" Then %>right<% Else %>left<% End If %>">
            		<% 
            		Select Case strScriptName 
            		Case "cart.asp"
            				If Session("PayCart") Then %>
            			<b><font size="1" face="Verdana" color="#000000">
			            <a class="LinkTitleExtra" href="javascript:Pay('payments/payCash.asp',476,130,'no')">
					<%=getagentTopLngStr("LtxtCash")%></a>&nbsp;&nbsp;
			            <a class="LinkTitleExtra" href="javascript:Pay('payments/payCheck.asp',686,300,'yes')">
					<%=getagentTopLngStr("LtxtCheck")%></a>&nbsp;&nbsp;
			            <a class="LinkTitleExtra" href="javascript:Pay('payments/payTrans.asp',554,220,'no')">
					<%=getagentTopLngStr("LtxtBankTrans")%></a>&nbsp;&nbsp;
			            <a class="LinkTitleExtra" href="javascript:Pay('payments/payCred.asp',640,400,'yes')">
					<%=getagentTopLngStr("LtxtCredCard")%></a></font></b>
		            <% End If
			        Case "agentclient.asp" %>
			            <b><font size="1" face="Verdana" color="#000000">
		            	<a class="LinkTitleExtra" href="javascript:Start('addCard/addresses.asp?AdresType=B',600,400,'yes')">
					<%=getagentTopLngStr("LtxtBillAdd")%></a>&nbsp;&nbsp;
		            	<a class="LinkTitleExtra" href="javascript:Start('addCard/addresses.asp?AdresType=S',600,400,'yes')">
					<%=getagentTopLngStr("LtxtShipAdd")%></a>&nbsp;&nbsp;
		            	<a class="LinkTitleExtra" href="javascript:Start('addCard/contacts.asp',600,400,'yes')">
					<%=getagentTopLngStr("LtxtContacts")%></a>
		            	</font></b>
		            	<% 
		            End Select %>
					</td>
				</tr>
				<% End If %>
			</table>
         </td>
		<td width="5">&nbsp;</td>
      </tr>
    </table>
    </td>
  </tr>
  <tr>
    <td colspan="2" background="ventas/images/art_ventas1_r4_c1.jpg" valign="top">
    <table border="0" cellpadding="0"  bordercolor="#111111" width="101%" style="border-collapse: collapse" cellspacing="0">
      <tr>
        <td width="117" height="215" background="ventas/images/<%=Session("rtl")%>art_conimg_r4_c1.jpg" valign="top">
    <div align="center"></div>
		</td>
      </tr>
      <tr>
        <td>
        <p>&nbsp;</p>
          <p>&nbsp;</p>
          <p>&nbsp;</p>
          <p>&nbsp;</p>
        </td>
      </tr>
      <tr>
        <td style="font-size: 4px" height="600">&nbsp;</td>
      </tr>
    </table>
    </td>
    <td colspan="11" background="ventas/images/<%=Session("rtl")%>backraya_1.jpg" valign="top"  style="background-repeat: repeat-y; background-color: white; <% If Session("rtl") = "rtl/" Then %>background-position: top right; <% End If %>">
    <table border="0" cellpadding="0"  bordercolor="#111111" width="100%" height="342">
      <tr>
        <td width="100%" valign="top" bgcolor="white">
        <div id="printThis">

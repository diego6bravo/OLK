<!--#include file="top.asp" -->
<!-- #INCLUDE file="FCKeditor/fckeditor.asp" -->
<!--#include file="lang/adminSecEdit.asp" -->
<!--#include file="adminTradSubmit.asp"-->
<% conn.execute("use [" & Session("olkdb") & "]") %>
<%
Dim oFCKeditor

sql = "select SelDes from OLKCommon"
set rs = conn.execute(sql)
SelDes = rs(0)
rs.close
%>
<script language="javascript">
var txtValSecNam = '<% If Request("UType") = "C" Then %><%=getadminSecEditLngStr("LtxtValSecNam")%><% Else %><%=getadminSecEditLngStr("LtxtValFormNam")%><% End If %>';
var txtValLink = '<%=getadminSecEditLngStr("LtxtValLink")%>';
var txtValRep = '<%=getadminSecEditLngStr("LtxtValRep")%>';
</script>
<script language="javascript" src="adminSec.js"></script>
<script language="javascript" src="js_up_down.js"></script>

<head>
<style type="text/css">
.style1 {
	text-align: center;
	background-color: #F7FBFF;
}
.style2 {
	background-color: #F7FBFF;
}
.style3 {
	direction: ltr;
}
.style4 {
				background-color: #E7F7FF;
				font-weight: bold;
}
</style>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>

<table border="0" cellpadding="0" width="100%" id="table3">
	<form method="POST" action="" name="frmAddEditSec">
	<% If Request("SecID") = "" Then %>
	<input type="hidden" name="SecNameTrad" value="<%=Server.HTMLEncode(Request("SecNameTrad"))%>">
	<input type="hidden" name="SecContentTrad" value="<%=Server.HTMLEncode(Request("SecContentTrad"))%>">
	<input type="hidden" name="SecSmallTextTrad" value="<%=Server.HTMLEncode(Request("SecSmallTextTrad"))%>">
	<% End If %>
	<tr>
		<td bgcolor="#E1F3FD"><b><font face="Verdana" size="1" color="#31659C">&nbsp;<% 
		Select Case Request("UType")
			Case "C" %>
				<% If Request("SecID") = "" Then %><%=getadminSecEditLngStr("LttlAddSec")%><% Else %><%=getadminSecEditLngStr("LttlEditSec")%><% End If %>
		<% 	Case "P", "A" %>
				<% If Request("SecID") = "" Then %><%=getadminSecEditLngStr("LttlAddForm")%><% Else %><%=getadminSecEditLngStr("LttlEditForm")%><% End If %>
		<% End Select %></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1" color="#4783C5">
		<% 
		Select Case Request("UType")
			Case "C" %><%=getadminSecEditLngStr("LttlSecNote")%>
		<%	Case "P", "A" %><%=getadminSecEditLngStr("LttlFormNote")%>
		<% End Select %></font></td>
	</tr>
	<% 
	If Request.Form.Count = 0 Then
		If Request("SecID") <> "" Then
			sql = "select SecName, SecOrder, Case Type When 'N' Then SecContent Else '' End SecContent, Case Type When 'L' Then SecContent Else '' End NewLink, Case Type When 'R' Then SecContent Else '' End rsIndex, IsNull(SecSmallText,'') SecSmallText, ReqLogin, Status, HideMainMenu, HideSecondMenu, Manual, " & _
			"Form, FormScript, FormQry, FormConfirmContent, SecContentEnableQry, SecContentQry, FormQryRS, FormQryLoop, ApplyCSS, Type " & _
			"from OLKSections where SecType = 'U' and SecID = " & Request("SecID")
			set rs = conn.execute(sql)
			NewName = rs("SecName")
			NewOrder = rs("SecOrder")
			ReqLogin = rs("ReqLogin")
			Status = rs("Status")
			NewContent = rs("SecContent")
			HideMainMenu = rs("HideMainMenu") = "Y"
			HideSecondMenu = rs("HideSecondMenu") = "Y"
			SecSmallText = rs("SecSmallText")
			Manual = rs("Manual")
			ApplyCSS = rs("ApplyCSS")
			Form = rs("Form")
			FormScript = rs("FormScript")
			FormQry = rs("FormQry")
			FormConfirmContent = rs("FormConfirmContent")
			SecContentEnableQry = rs("SecContentEnableQry") = "Y"
			SecContentQry = rs("SecContentQry")
			FormQryRS = rs("FormQryRS") = "Y"
			FormQryLoop = rs("FormQryLoop")
			EditType = rs("Type")
			NewLink = rs("NewLink")
			rsIndex = rs("rsIndex")
		Else
			NewName = ""
			NewOrder = Request("NewOrder")
			ReqLogin = "N"
			Status = "N"
			HideMainMenu = False
			HideSecondMenu = True
			NewContent = "<font face=""verdana"" size=""1""><div></div></font>"
			SecSmallText = ""
			Manual = "N"
			ApplyCSS = "Y"
			If Request("UType") = "P" Then Form = "Y"
			SecContentEnableQry = False
			FormQryRS = False
			EditType = "N"
			NewLink = ""
			rsIndex = ""
		End If
	Else
		NewName = Request("NewName")
		NewOrder = Request("NewOrder")
		
		If Request("NewReqLogin") = "Y" Then ReqLogin = "Y" Else ReqLogin = "N"
		If Request("NewActive") = "Y" Then Status = "A" Else Status = "N"
		HideMainMenu = Request("HideMainMenu") = "Y"
		HideSecondMenu = Request("HideSecondMenu") = "Y" 
		If Request("NewManual") = "Y" Then Manual = "Y" Else Manual = "N"
		If Request("ApplyCSS") = "Y" Then ApplyCSS = "Y" Else ApplyCSS = "N"
		If Request("Form") = "Y" Then Form = "Y" Else Form = "N"
		SecContentEnableQry = Request("SecContentEnableQry") = "Y"
		FormQryRS = Request("FormQryRS") = "Y"

		FormScript = Request("FormScript")
		FormConfirmContent = Request("FormConfirmContent")
		FormQry = Request("FormQry")
		SecContentQry = Request("SecContentQry")
		FormQryLoop = Request("FormQryLoop")
		
		NewContent = Request("NewContent")
		SecSmallText = Request("SecSmallText")
		EditType = Request("Type")
		
		NewLink = Request("NewLink")
		rsIndex = Request("rsIndex")
	End If %>
	<script language="javascript">
	var isManual = '<% If EditType = "N" Then %><%=Manual%><% End If %>';
	</script>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr>
				<td class="style4">
				<font face="Verdana" size="1" color="#31659C"><%=getadminSecEditLngStr("DtxtType")%></font></td>
				<td class="style2">
				<select size="1" name="Type" onchange="reloadEdit();">
				<option value="N"><%=getadminSecEditLngStr("DtxtContent")%></option>
				<option <% If EditType = "R" Then %>selected<% End If %> value="R"><%=getadminSecEditLngStr("DtxtReport")%></option>
				<option <% If EditType = "L" Then %>selected<% End If %> value="L"><%=getadminSecEditLngStr("DtxtLink")%></option>
				</select></td>
			</tr>
			<tr>
				<td class="style4">
				<font face="Verdana" size="1" color="#31659C"><%=getadminSecEditLngStr("DtxtName")%></font></td>
				<td class="style2">
				<table cellpadding="0" cellspacing="0" border="0" style="width: 100%;">
					<tr>
						<td><input class="input" type="text" name="NewName" size="44" style="width: 100%;" onkeydown="return chkMax(event, this, 100);" value="<%=Server.HTMLEncode(NewName)%>"></td>
						<td style="width: 16px;"><a href="javascript:doFldTrad('Sections', 'SecType,SecID', 'U,<%=Request("SecID")%>', 'AlterSecName', 'T', <% If Request("SecID") <> "" Then %>null<% Else %>document.frmAddEditSec.SecNameTrad<% End If %>);"><img src="images/trad.gif" alt="<%=getadminSecEditLngStr("DtxtTranslate")%>" border="0"></a></td>
					</tr>
				</table>
				</td>
			</tr>
			<% If EditType = "L" Then %>
			<tr>
				<td class="style4">
				<font face="Verdana" size="1" color="#31659C"><%=getadminSecEditLngStr("DtxtLink")%></font></td>
				<td class="style2">
				<input class="input" type="text" name="NewLink" size="44" onkeydown="return chkMax(event, this, 100);" value="<%=Server.HTMLEncode(NewLink)%>">
				</td>
			</tr>
			<% Else %>
			<input type="hidden" name="NewLink" value="<%=Server.HTMLEncode(NewLink)%>">
			<% End If %>
			<% If EditType = "R" Then %>
			<tr>
				<td class="style4">
				<font face="Verdana" size="1" color="#31659C"><%=getadminSecEditLngStr("DtxtReport")%></font></td>
				<td class="style2">
				<select size="1" name="rsIndex">
				<option></option>
				<% 
				LastRG = ""
				Select Case Request("UType")
					Case "C"
						UserType = "C"
					Case "P", "A"
						UserType = "V"
				End Select
				sql = "select T1.rgName, T0.rsIndex, T0.rsName " & _
						"from OLKRS T0 " & _
						"inner join OLKRG T1 on T1.rgIndex = T0.rgIndex " & _
						"where T1.UserType = '" & UserType & "' and T0.Active = 'Y' " & _
						"order by T1.rgName, T0.rsName "
				set rs = conn.execute(sql)
				do while not rs.eof
				If LastRG <> rs("rgName") Then
					If LastRG <> "" Then Response.Write "</optgroup>"
					Response.WRite "<optgroup label=""" & myHTMLEncode(rs("rgName")) & """>"
					LastRG = rs("rgName")
				End If %>
				<option <% If rsIndex = CStr(rs("rsIndex")) Then %>selected<% End If %> value="<%=rs("rsIndex")%>"><%=myHTMLEncode(rs("rsName"))%></option>
				<% rs.movenext
				loop
				Response.Write "</optgroup>" %>
				</select></td>
			</tr>
			<% Else %>
			<input type="hidden" name="rsIndex" value="<%=rsIndex%>">
			<% End If %>
			<tr>
				<td class="style4">
				<b>
				<font face="Verdana" size="1" color="#31659C"><%=getadminSecEditLngStr("DtxtOrder")%></font></b></td>
				<td class="style2">				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td>
						<input id="NewOrder" name="NewOrder" class="input" value="<%=NewOrder%>" size="7" onfocus="this.select();" onmouseup="event.preventDefault()" onkeydown="return chkMax(event, this, 6);"></td>
						<td valign="middle">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td><img src="images/img_nud_up.gif" id="btnNewOrderUp"></td>
							</tr>
							<tr>
								<td><img src="images/spacer.gif"></td>
							</tr>
							<tr>
								<td><img src="images/img_nud_down.gif" id="btnNewOrderDown"></td>
							</tr>
						</table>
						</td>
					</tr>
				</table>
				<script language="javascript">NumUDAttach('frmAddEditSec', 'NewOrder', 'btnNewOrderUp', 'btnNewOrderDown');</script></td>
			</tr>
			<% If EditType = "N" Then %>
			<tr>
				<td class="style4">
				&nbsp;</td>
				<td class="style2">
				<input type="checkbox" name="Form" onclick="reloadEdit();" id="Form" <% If Form = "Y" Then Response.Write "checked" %> value="Y" style="border-style:solid; border-width:0; background:background-image"><b><font face="Verdana" size="1" color="#31659C"><label for="Form"><%=getadminSecEditLngStr("DtxtForm")%></label></font></b></td>
			</tr>
			<% Else %>
			<input type="hidden" name="Form" value="<%=Request("Form")%>">
			<% End If %>
			<tr>
				<td class="style4">
				&nbsp;</td>
				<% 
				Checked = ""
				Disabled = ""
				Select Case EditType
					Case "R"
						If Request("UType") = "C" Then 
							Disabled = "disabled "
							Checked = "checked "
						Else
							If ReqLogin = "Y" Then Checked = "checked "
						End If
					Case "N", "L"
						If ReqLogin = "Y" Then Checked = "checked "
				End Select %>
				<td class="style2">
				<span <%=Disabled%>>
				<input type="checkbox" name="NewReqLogin" id="NewReqLogin" <%=Checked%><%=Disabled%> value="Y" style="border-style:solid; border-width:0; background:background-image"><b><font face="Verdana" size="1" color="#31659C"><label for="NewReqLogin"><%=getadminSecEditLngStr("LtxtSesion")%></label></font></b></span></td>
			</tr>
			<tr>
				<td class="style4">
				&nbsp;</td>
				<td class="style2">
				<b>
				<font face="Verdana" size="1" color="#31659C">
				<input type="checkbox" name="HideMainMenu" id="HideMainMenu" <% If HideMainMenu Then %>checked<% End If %> value="Y" style="border-style:solid; border-width:0; background:background-image"><label for="HideMainMenu"><%=getadminSecEditLngStr("LtxtHideMnu")%></label></font></b></td>
			</tr>
			<% If Request("UType") = "C" Then %>
			<tr>
				<td class="style4">
				&nbsp;</td>
				<td class="style2">
				<input type="checkbox" name="HideSecondMenu" id="HideSecondMenu" <% If HideSecondMenu Then %>checked<% End If %> value="Y" style="border-style:solid; border-width:0; background:background-image"><b><font face="Verdana" size="1" color="#31659C"><label for="HideSecondMenu"><%=getadminSecEditLngStr("LtxtHideSecMenu")%></label></font></b></td>
			</tr>
			<% End If %>
			<% If EditType = "N" Then %>
			<tr>
				<td class="style4">
				&nbsp;</td>
				<td class="style2">
				<input type="checkbox" onclick="reloadEdit();" name="NewManual" id="NewManual" <% If Manual = "Y" Then Response.Write "checked" %> value="Y" style="border-style:solid; border-width:0; background:background-image"><b><font face="Verdana" size="1" color="#31659C"><label for="NewManual"><%=getadminSecEditLngStr("LtxtManual")%></label></font></b></td>
			</tr>
			<tr>
				<td class="style4">
				&nbsp;</td>
				<td class="style2">
				<input type="checkbox" name="ApplyCSS" id="ApplyCSS" value="Y" <% If ApplyCSS = "Y" Then %>checked<% End If %>  style="border-style:solid; border-width:0; background:background-image"><b><font face="Verdana" size="1" color="#31659C"><label for="ApplyCSS"><%=getadminSecEditLngStr("LtxtApplyCSS")%></label></font></b></td>
			</tr>
			<% Else %>
			<input type="hidden" name="NewManual" value="<%=Request("NewManual")%>">
			<input type="hidden" name="ApplyCSS" value="<%=Request("ApplyCSS")%>">
			<% End If %>
			<tr>
				<td class="style4">
				&nbsp;</td>
				<td class="style2">
				<input type="checkbox" name="NewActive" id="NewActive" <% If Status = "A" Then %>checked<% End If %> value="Y" style="border-style:solid; border-width:0; background:background-image"><b><font face="Verdana" size="1" color="#31659C"><label for="NewActive"><%=getadminSecEditLngStr("DtxtActive")%></label></font></b></td>
			</tr>
			<% If Request("UType") = "C" Then %>
			<tr bgcolor="#E2F3FC">
				<td colspan="2">
				<p align="center"><b>
				<font face="Verdana" size="1" color="#31659C"><%=getadminSecEditLngStr("LtxtShortText")%></font></b></td>
			</tr>
			<tr bgcolor="#F7FBFF">
				<td colspan="2">
				<p align="center">
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td><iframe name="testSmallText" src="clear.asp" target="_blank">Your browser does not support inline frames or is currently configured not to display inline frames.</iframe>
						</td>
						<td width="16" valign="bottom"><a href="javascript:doFldTrad('Sections', 'SecType,SecID', 'U,<%=Request("SecID")%>', 'AlterSecSmallText', 'R', <% If Request("SecID") <> "" Then %>null<% Else %>document.frmAddEditSec.SecSmallTextTrad<% End If %>);">
						<img src="images/trad.gif" alt="<%=getadminSecEditLngStr("DtxtTranslate")%>" border="0"></a></td>
					</tr>
				</table>
				<script language="javascript">
				var tblTop = '<table cellpadding=\"0\" cellspacing=\"0\" border=\"0\" width=\"100%\"><tr><td class=\"TblWhiteHome\">';
				var tblBtm = '</td></tr></table>';
				function loadSmallText()
				{
					testSmallText.changeStyle('design/<%=SelDes%>/style/stylenuevo.css');
					testSmallText.document.body.innerHTML = '';
					testSmallText.document.body.innerHTML = tblTop + '<%=Replace(myHTMLEncode(SecSmallText), VbNewLine, "\n") %>' + tblBtm;
					document.frmAddEditSec.SecSmallText.value = '<%=Replace(myHTMLEncode(SecSmallText), VbNewLine, "\n") %>';
				}
				function doEditSmallText()
				{
					OpenWin = window.open('adminSecEditSmallText.asp', 'OpenWin', 'toolbar=no,menubar=no,location=no,scrollbars=no,resizable=no, width=660,height=300');
					document.frmSmallText.smallText.value = document.frmAddEditSec.SecSmallText.value;
					document.frmSmallText.submit();
				}
				function setSmallText(value)
				{
					document.frmAddEditSec.SecSmallText.value = value;
					testSmallText.document.body.innerHTML = '';
					testSmallText.document.body.innerHTML = tblTop + value + tblBtm;
				}
				</script>
				<br>
				<input type="button" value="<%=getadminSecEditLngStr("DtxtEdit")%>" name="btnEditSText" class="OlkBtn" onclick="doEditSmallText();"></td>
			</tr>
			<% End If %>
			<input type="hidden" name="SecSmallText" value="">
			<% If EditType = "N" Then %>
			<% If Form = "Y" Then %>
			<tr>
				<td colspan="2">
				<table border="0" cellpadding="0" width="300" id="table24">
					<tr>
						<td bgcolor="#D9F5FF" align="center" style="border: 1px solid #31659C" onclick="javascript:showTab('tabCont');" onmouseover="javascript:if(document.getElementById('trContent').style.display=='none')this.bgColor='#D9F5FF';" onmouseout="javascript:if(document.getElementById('trContent').style.display=='none')this.bgColor='#BFEEFE';" id="btnTabCont">
				<b><font color="#31659C" face="Verdana" size="1"><%=getadminSecEditLngStr("LtxtContent")%></font></b></td>
						<td bgcolor="#BFEEFE" align="center" style="border: 1px solid #31659C; cursor: hand" onclick="javascript:showTab('tabForm');" onmouseover="javascript:if(document.getElementById('tblForm').style.display=='none')this.bgColor='#D9F5FF';" onmouseout="javascript:if(document.getElementById('tblForm').style.display=='none')this.bgColor='#BFEEFE';" id="btnTabForm">
				<b>
				<font color="#31659C" face="Verdana" size="1"><%=getadminSecEditLngStr("DtxtForm")%></font></b></td>
					</tr>
				</table>
				</td>
			</tr>
			<% Else %>
			<tr bgcolor="#E2F3FC">
				<td colspan="2">
				<p align="center"><b>
				<font face="Verdana" size="1" color="#31659C"><%=getadminSecEditLngStr("LtxtContent")%></font></b></td>
			</tr>
			<% End If %>
			<tr id="trContent">
				<td colspan="2">
				<p align="center">
				<table cellpadding="0" cellspacing="0" border="0" width="100%">
					<% If Manual = "Y" Then %>
					<tr id="trEditPrev">
						<td>
							<table border="0" id="table13" cellpadding="0">
								<tr>
									<td id="btnPrevContent" bgcolor="#D9F5FF" align="center" style="border: 1px solid #31659C; padding-left: 10px; padding-right: 10px;" onclick="showPrevContent();" onmouseover="javascript:if(document.getElementById('PrevContent').style.display=='none')this.bgColor='#D9F5FF';" onmouseout="javascript:if(document.getElementById('PrevContent').style.display=='none')this.bgColor='#BFEEFE';">
									<font color="#31659C" face="Verdana" size="1"><b><%=getadminSecEditLngStr("DtxtPreview")%></b></font></td>
									<td id="btnEditContent" bgcolor="#BFEEFE" align="center" style="border: 1px solid #31659C; cursor: hand; padding-left: 10px; padding-right: 10px;" onclick="showEditContent();" onmouseover="javascript:if(document.frmAddEditSec.NewContent.style.display=='none')this.bgColor='#D9F5FF';" onmouseout="javascript:if(document.frmAddEditSec.NewContent.style.display=='none')this.bgColor='#BFEEFE';">
									<font color="#31659C" face="Verdana" size="1"><b><%=getadminSecEditLngStr("DtxtEdit")%></b></font></td>
									<td id="btnEditRS" bgcolor="#BFEEFE" align="center" style="border: 1px solid #31659C; cursor: hand; padding-left: 10px; padding-right: 10px;" onclick="showRSContent();" onmouseover="javascript:if(document.getElementById('frmSectionsRS').style.display=='none')this.bgColor='#D9F5FF';" onmouseout="javascript:if(document.getElementById('frmSectionsRS').style.display=='none')this.bgColor='#BFEEFE';">
									<font color="#31659C" face="Verdana" size="1"><b><%=getadminSecEditLngStr("LtxtRecordSets")%></b></font></td>
									<% If Form = "Y" Then %><td id="btnEditQuery" bgcolor="#BFEEFE" align="center" style="border: 1px solid #31659C; cursor: hand; padding-left: 10px; padding-right: 10px;" onclick="showEditQuery();" onmouseover="javascript:if(document.getElementById('frmQuery').style.display=='none')this.bgColor='#D9F5FF';" onmouseout="javascript:if(document.getElementById('frmQuery').style.display=='none')this.bgColor='#BFEEFE';" class="style3">
									<font color="#31659C" face="Verdana" size="1"><b><%=getadminSecEditLngStr("DtxtQuery")%></b></font></td>
									<td id="btnEditScript" bgcolor="#BFEEFE" align="center" style="border: 1px solid #31659C; cursor: hand; padding-left: 10px; padding-right: 10px;" onclick="showEditScript();" onmouseover="javascript:if(document.getElementById('frmScript').style.display=='none')this.bgColor='#D9F5FF';" onmouseout="javascript:if(document.getElementById('frmScript').style.display=='none')this.bgColor='#BFEEFE';">
									<font color="#31659C" face="Verdana" size="1"><b><%=getadminSecEditLngStr("DtxtScript")%></b></font></td><% End If %>
								</tr>
							</table>
						</td>
					</tr>
					<% End If %>
					<tr bgcolor="#F7FBFF">
						<td><% If Manual = "N" Then %>
								<% Else %>
								<iframe name="PrevContent" id="PrevContent" src="clear.asp" target="_blank" style="width: 100%; height: 560px">
								Your browser does not support inline frames or is currently 
								configured not to display inline frames.
								</iframe>
								<% End If %>
								<%
								If Manual <> "Y" Then
									Set oFCKeditor = New FCKeditor
									oFCKeditor.BasePath = "FCKeditor/"
									oFCKeditor.Height = 560
									oFCKEditor.ToolbarSet = "Default"
									oFCKEditor.Value = myHTMLEncode(NewContent)
									oFCKEditor.Config("AutoDetectLanguage") = False
									If Session("myLng") <> "pt" Then
										oFCKEditor.Config("DefaultLanguage") = Session("myLng")
									Else
										oFCKEditor.Config("DefaultLanguage") = "pt-br"
									End If
									oFCKeditor.Create "NewContent"
								Else
								%>
								<textarea onkeydown="return catchTab(this,event)" rows="35" id="NewContent" name="NewContent" dir="ltr" style="width: 100%; <% If Manual = "Y" Then %>display: none<% End If %>" cols="1"><%=Server.HTMLEncode(NewContent)%></textarea>
								<% End If %>
						</td>
						<td width="16" valign="bottom" id="trNewContentTrans">
						<a href="javascript:doFldTrad('Sections', 'SecType,SecID', 'U,<%=Request("SecID")%>', 'AlterSecContent', '<% If Manual = "Y" Then %>M<% Else %>R<% End If %>', <% If Request("SecID") <> "" Then %>null<% Else %>document.frmAddEditSec.SecContentTrad<% End If %>);">
						<img src="images/trad.gif" alt="<%=getadminSecEditLngStr("DtxtTranslate")%>" border="0"></a></td>
					</tr>
					<% If Manual = "Y" Then %>
					<tr id="frmSectionsRS" style="display: none">
						<td bgcolor="#F7FBFF">
							<iframe src="adminSecRS.asp?SecID=<%=Request("SecID")%>" style="width: 100%; height: 560px">
							Your browser does not support inline frames or is currently 
							configured not to display inline frames.
							</iframe>
						</td>
					</tr>
					<% End If %>
					<% If Form = "Y" Then %>
					<tr id="frmScript" style="display: none; ">
						<td bgcolor="#F7FBFF">
						<table cellpadding="0" cellspacing="0" border="0" width="100%">
							<tr>
								<td><textarea onkeydown="return catchTab(this,event)" rows="25" name="FormScript" cols="10" dir="ltr" class="input" style="height:560px; width: 100%;"><%=myHTMLEncode(FormScript)%></textarea>
								</td>
								<% If Request("SecID") <> "" Then %>
								<td width="16" valign="bottom">
								<a href="javascript:doFldTrad('Sections', 'SecType,SecID', 'U,<%=Request("SecID")%>', 'AlterFormScript', 'M', null);">
								<img src="images/trad.gif" alt="<%=getadminSecEditLngStr("DtxtTranslate")%>" border="0"></a>
								</td>
								<% End If %>
							</tr>
						</table>
						</td>
					</tr>
					<tr id="frmQuery" style="display: none; ">
						<td>
						<table cellpadding="0" cellspacing="0" border="0" width="100%">
							<tr>
								<td bgcolor="#F7FBFF"><font face="Verdana" size="1" color="#4783C5"><input type="checkbox" style="border-style:solid; border-width:0; background:background-image" name="SecContentEnableQry" id="chkSecContentEnableQry" value="Y" <% If SecContentEnableQry Then %>checked<% End If %>>
								<label for="chkSecContentEnableQry"><%=getadminSecEditLngStr("LtxtSecContentEnableQ")%></label></font>
								</td>
							</tr>
							<tr>
								<td bgcolor="#F7FBFF"><textarea onkeydown="return catchTab(this,event)" rows="25" name="SecContentQry" cols="10" class="input" style="height:560px; width: 100%;" dir="ltr"><%=myHTMLEncode(SecContentQry)%></textarea>
								</td>
							</tr>
						</table>
						</td>
					</tr>
					<% End If %>
				</table>				
				</td>
			</tr>
			<% If Form = "Y" Then %>
			<tr>
				<td colspan="2">
				<table border="0" cellpadding="0" width="100%" id="tblForm" style="display: none;">
					<tr>
						<td bgcolor="#E2F3FC">&nbsp;</td>
						<td bgcolor="#F7FBFF"><font face="Verdana" size="1" color="#4783C5"><input type="checkbox" style="border-style:solid; border-width:0; background:background-image" name="FormQryRS" id="chkFormQryRS" value="Y" <% If FormQryRS Then %>checked<% End If %>>
						<label for="chkFormQryRS"><%=getadminSecEditLngStr("LtxtEnableRS")%></label></font></td>
					</tr>
					<tr>
						<td width="120" valign="top" bgcolor="#E2F3FC"><b>
						<font face="Verdana" size="1" color="#31659C"><%=getadminSecEditLngStr("DtxtQuery")%></font></b></td>
						<td valign="bottom" bgcolor="#F7FBFF">
						<table border="0" cellpadding="0" cellspacing="1" width="560">
							<tr>
								<td>
								<textarea onkeydown="return catchTab(this,event)" rows="6" name="FormQry" dir="ltr" cols="10" class="input" style="width: 100%;"<% If 1 = 2 Then %> onkeypress="javascript:btnVerfy.disabled = false;"<% End If %>><%=myHTMLEncode(FormQry)%></textarea></td>
								<td valign="bottom" width="24"><% If 1 = 2 Then %><input type="button" value="a" disabled name="btnVerfy" id="btnVerfy" style="font-family: Webdings; border: 1px solid; font-size:12pt; width:24; height:22; font-weight:bold"><% End If %>&nbsp;</td>
							</tr>
						</table>
						</td>
						</tr>
					<tr>
						<td width="120" valign="top" bgcolor="#E2F3FC"><b>
						<font face="Verdana" size="1" color="#31659C"><%=getadminSecEditLngStr("DtxtQuery")%> (<%=getadminSecEditLngStr("LtxtLoop")%>)</font></b></td>
						<td valign="bottom" bgcolor="#F7FBFF">
						<table border="0" cellpadding="0" cellspacing="1" width="560">
							<tr>
								<td>
								<textarea onkeydown="return catchTab(this,event)" rows="6" name="FormQryLoop" cols="10" class="input" style="width: 100%; "<% If 1 = 2 Then %> onkeypress="javascript:btnVerfy.disabled = false;"<% End If %> dir="ltr"><%=myHTMLEncode(FormQryLoop)%></textarea></td>
								<td valign="bottom" width="24"><% If 1 = 2 Then %><input type="button" value="a" disabled name="btnVerfy" id="btnVerfy" style="font-family: Webdings; border: 1px solid; font-size:12pt; width:24; height:22; font-weight:bold"><% End If %>&nbsp;</td>
							</tr>
							<tr>
								<td>
								<font size="1" color="#4783C5" face="Verdana">
								<b><font face="Verdana" size="1" color="#31659C">
																<%=getadminSecEditLngStr("DtxtVariables")%><br>
								</font></b><span id="txtVars">
								<span dir="ltr">@CardCode</span> = <%=getadminSecEditLngStr("DtxtClientCode")%></span></font></td>
								<td valign="bottom" width="24">&nbsp;</td>
							</tr>
						</table>
						</td>
						</tr>
						<tr>
							<td width="120" bgcolor="#E2F3FC" valign="top"><b>
							<font face="Verdana" size="1" color="#31659C"><%=getadminSecEditLngStr("LtxtConfCont")%></font></b></td>
							<td bgcolor="#F7FBFF">
							<table cellpadding="0" cellspacing="0" border="0" width="100%">
							<tr>
								<td>
								<%
								If Manual = "N" Then
									Set oFCKeditor = New FCKeditor
									oFCKeditor.BasePath = "FCKeditor/"
									oFCKeditor.Height = 300
									oFCKEditor.ToolbarSet = "Custom"
									oFCKEditor.Value = myHTMLEncode(FormConfirmContent)
									oFCKEditor.Config("AutoDetectLanguage") = False
									If Session("myLng") <> "pt" Then
										oFCKEditor.Config("DefaultLanguage") = Session("myLng")
									Else
										oFCKEditor.Config("DefaultLanguage") = "pt-br"
									End If
									oFCKeditor.Create "FormConfirmContent"
								Else %><textarea onkeydown="return catchTab(this,event)" rows="35" id="FormConfirmContent" name="FormConfirmContent" dir="ltr" style="width: 100%; " cols="1"><%=myHTMLEncode(FormConfirmContent)%></textarea>
								<% End If %>
								</td>
								<% If Request("SecID") <> "" Then %>
								<td width="16" valign="bottom"><a href="javascript:doFldTrad('Sections', 'SecType,SecID', 'U,<%=Request("SecID")%>', 'AlterFormConfirmContent', 'R', null);">
								<img src="images/trad.gif" alt="<%=getadminSecEditLngStr("DtxtTranslate")%>" border="0"></a></td>
								<% End If %>
							</tr>
						</table>
						</td>
						</tr>
					</table>
				</td>
			</tr>
			<% End If %>
			<% Else %>
			<input type="hidden" name="NewContent" value="<%=Server.HTMLEncode(Request("NewContent"))%>">
			<input type="hidden" name="FormScript" value="<%=Server.HTMLEncode(Request("FormScript"))%>">
			<input type="hidden" name="SecContentEnableQry" value="<%=Server.HTMLEncode(Request("SecContentEnableQry"))%>">
			<input type="hidden" name="SecContentQry" value="<%=Server.HTMLEncode(Request("SecContentQry"))%>">
			<input type="hidden" name="FormQryRS" value="<%=Server.HTMLEncode(Request("FormQryRS"))%>">
			<input type="hidden" name="FormQry" value="<%=Server.HTMLEncode(Request("FormQry"))%>">
			<input type="hidden" name="FormQryLoop" value="<%=Server.HTMLEncode(Request("FormQryLoop"))%>">
			<% End If %>
			</table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table5">
			<tr>
				<td width="77">
				<input type="submit" value="<%=getadminSecEditLngStr("DtxtApply")%>" name="btnApply" class="OlkBtn" onclick="return valFrmEdit();"></td>
				<td width="77">
				<input type="submit" value="<%=getadminSecEditLngStr("DtxtSave")%>" name="btnSave" class="OlkBtn" onclick="return valFrmEdit();"></td>
				<td><hr color="#0D85C6" size="1"></td>
				<td width="77">
				<input type="button" value="<%=getadminSecEditLngStr("DtxtCancel")%>" name="B1" class="OlkBtn" onclick="javascript:if(confirm('<%=getadminSecEditLngStr("DtxtConfCancel")%>'))window.location.href='adminSec.asp?UType=<%=Request("UType")%>'"></td>
			</tr>
		</table>
		</td>
	</tr>
	<input type="hidden" name="submitCmd" value="adminSec">
	<input type="hidden" name="uCmd" value="edit">
	<input type="hidden" name="SecID" value="<%=Request("SecID")%>">
	<input type="hidden" name="rCount" value="<%=Request("rCount")%>">
	<input type="hidden" name="UType" value="<%=Request("UType")%>">
	</form>
	<tr>
		<td>&nbsp;</td>
	</tr>
</table>
<form name="frmSmallText" action="adminSecEditSmallText.asp" method="post" target="OpenWin">
<input type="hidden" name="smallText" value="">
</form>
<% If EditType = "N" Then %>
<script language="javascript">
<% If Form = "Y" Then %>
function showTab(tab)
{
	switch (tab)
	{
		case 'tabCont':
			document.getElementById('btnTabCont').style.backgroundColor = '#D9F5FF';
			document.getElementById('btnTabCont').style.cursor = '';
			document.getElementById('btnTabForm').style.backgroundColor = '#BFEEFE';
			document.getElementById('btnTabForm').style.cursor = 'hand';
			document.getElementById('tblForm').style.display = 'none';
			document.getElementById('trContent').style.display = '';
			<% If Manual = "Y" Then %>
			document.getElementById('trEditPrev').style.display = '';
			<% End If %>
			
			break;
		case 'tabForm':
			document.getElementById('btnTabCont').style.backgroundColor = '#BFEEFE';
			document.getElementById('btnTabCont').style.cursor = 'hand';
			document.getElementById('btnTabForm').style.backgroundColor = '#D9F5FF';
			document.getElementById('btnTabForm').style.cursor = '';
			document.getElementById('tblForm').style.display = '';
			document.getElementById('trContent').style.display = 'none';
			<% If Manual = "Y" Then %>
			document.getElementById('trEditPrev').style.display = 'none';
			<% End If %>
			break;
	}
}
<% End If %>
<% If Manual = "Y" Then %>
function newPreview()
{
	strPrev = '<div class="TblGeneral">' + document.frmAddEditSec.NewContent.value + '</div>';
	//strPrev = strPrev.replace('src="', 'src=\"')
	strPrev = strPrev.replace('{dbName}', '<%=Session("olkdb")%>');
	strPrev = strPrev.replace('{SelDes}', '<%=SelDes%>');
	strPrev = strPrev.replace('<A href', '<A class=""LinkTop"" href');
	strPrev = strPrev.replace('<a href', '<A class=""LinkTop"" href');
	
	doPreview(strPrev);
}

function showPrevContent()
{
	if (document.getElementById('PrevContent').style.display == 'none')
	{
		newPreview();
		
		document.getElementById('PrevContent').style.display = '';
		document.getElementById('NewContent').style.display = 'none';
		document.getElementById('frmSectionsRS').style.display = 'none';

		btnPrevContent.bgColor = '#D9F5FF';
		btnPrevContent.style.cursor = '';
		btnEditContent.bgColor = '#BFEEFE';
		btnEditContent.style.cursor = 'hand';
		btnEditRS.bgColor = '#BFEEFE';
		btnEditRS.style.cursor = 'hand';
		
		trNewContentTrans.style.display = '';

		<% If Form = "Y" Then %>
		document.getElementById('frmQuery').style.display = 'none';
		
		btnEditQuery.bgColor = '#BFEEFE';
		btnEditQuery.style.cursor = 'hand';
		
		btnEditScript.bgColor = '#BFEEFE';
		btnEditScript.style.cursor = 'hand';
		document.getElementById('frmScript').style.display = 'none';
		<% End If %>
	}
}
function showEditContent()
{
	if (document.getElementById('NewContent').style.display == 'none')
	{
		document.getElementById('PrevContent').style.display = 'none';
		document.getElementById('NewContent').style.display = '';
		document.getElementById('frmSectionsRS').style.display = 'none';
		
		btnPrevContent.bgColor = '#BFEEFE';
		btnPrevContent.style.cursor = 'hand';
		btnEditContent.bgColor = '#D9F5FF';
		btnEditContent.style.cursor = '';
		btnEditRS.bgColor = '#BFEEFE';
		btnEditRS.style.cursor = 'hand';

		trNewContentTrans.style.display = '';

		<% If Form = "Y" Then %>
		document.getElementById('frmQuery').style.display = 'none';
		
		btnEditQuery.bgColor = '#BFEEFE';
		btnEditQuery.style.cursor = 'hand';
		
		btnEditScript.bgColor = '#BFEEFE';
		btnEditScript.style.cursor = 'hand';
		document.getElementById('frmScript').style.display = 'none';
		<% End If %>
	}
}
function showRSContent()
{
	if (document.getElementById('frmSectionsRS').style.display == 'none')
	{
		document.getElementById('PrevContent').style.display = 'none';
		document.getElementById('NewContent').style.display = 'none';
		document.getElementById('frmSectionsRS').style.display = '';
		
		btnPrevContent.bgColor = '#BFEEFE';
		btnPrevContent.style.cursor = 'hand';
		btnEditContent.bgColor = '#BFEEFE';
		btnEditContent.style.cursor = 'hand';
		btnEditRS.bgColor = '#D9F5FF';
		btnEditRS.style.cursor = '';

		trNewContentTrans.style.display = 'none';

		<% If Form = "Y" Then %>
		document.getElementById('frmQuery').style.display = 'none';
		
		btnEditQuery.bgColor = '#BFEEFE';
		btnEditQuery.style.cursor = 'hand';
		
		btnEditScript.bgColor = '#BFEEFE';
		btnEditScript.style.cursor = 'hand';
		document.getElementById('frmScript').style.display = 'none';
		<% End If %>
	}
}
function showEditQuery()
{
	if (document.getElementById('frmQuery').style.display == 'none')
	{
		document.getElementById('PrevContent').style.display = 'none';
		document.getElementById('NewContent').style.display = 'none';
		document.getElementById('frmQuery').style.display = '';
		document.getElementById('frmSectionsRS').style.display = 'none';
		
		btnPrevContent.bgColor = '#BFEEFE';
		btnPrevContent.style.cursor = 'hand';
		btnEditContent.bgColor = '#BFEEFE';
		btnEditContent.style.cursor = 'hand';
		btnEditRS.bgColor = '#BFEEFE';
		btnEditRS.style.cursor = 'hand';

		trNewContentTrans.style.display = 'none';

		btnEditQuery.bgColor = '#D9F5FF';
		btnEditQuery.style.cursor = '';
		
		<% If Form = "Y" Then %>
		btnEditScript.bgColor = '#BFEEFE';
		btnEditScript.style.cursor = 'hand';
		document.getElementById('frmScript').style.display = 'none';
		<% End If %>
	}
}
<% If Form = "Y" Then %>
function showEditScript()
{
	if (document.getElementById('frmScript').style.display == 'none')
	{
		document.getElementById('PrevContent').style.display = 'none';
		document.getElementById('NewContent').style.display = 'none';
		document.getElementById('frmQuery').style.display = 'none';
		document.getElementById('frmScript').style.display = '';
		document.getElementById('frmSectionsRS').style.display = 'none';
		
		btnPrevContent.bgColor = '#BFEEFE';
		btnPrevContent.style.cursor = 'hand';
		btnEditContent.bgColor = '#BFEEFE';
		btnEditContent.style.cursor = 'hand';
		btnEditQuery.bgColor = '#BFEEFE';
		btnEditQuery.style.cursor = 'hand';
		btnEditRS.bgColor = '#BFEEFE';
		btnEditRS.style.cursor = 'hand';

		trNewContentTrans.style.display = 'none';

		btnEditScript.bgColor = '#D9F5FF';
		btnEditScript.style.cursor = '';
	}
}
<% End If %>
function doPreview(Content)
{
	PrevContent.document.body.innerHTML = Content;
}

function doLoadPreview()
{
	PrevContent.changeStyle('design/<%=SelDes%>/style/stylenuevo.css');
	newPreview();
}
<% End If %>
</script>
<% End If %><!--#include file="bottom.asp" -->
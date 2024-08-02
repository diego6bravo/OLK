<% pageTtl = "Editar Reporte"
pDesc = pageTtl
pCod = "edit" %>
<!--#include file="repTop.inc" -->
<!--#include file="lang/adminRepEdit.asp" -->
<!--#include file="adminTradSubmit.asp"-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<script type="text/javascript" src="scr/calendar.js"></script>
<script type="text/javascript" src="scr/lang/calendar-<%=Left(Session("myLng"), 2)%>.js"></script>
<script type="text/javascript" src="scr/calendar-setup.js"></script>
<% 
set rs = Server.CreateObject("ADODB.recordset")
SQL = "Select rsIndex, rsName, rsDesc, rsQuery, rsTop, rsTopDef, rgIndex, Active, Refresh, LinkOnly, (select UserType from OLKRG where rgIndex = T0.rgIndex) UserType FROM OLKRS T0 where rsIndex = " & Request("rsIndex")
SET RS = Conn.Execute(SQL)
rsIndex = rs("rsIndex")
UserType = rs("UserType")
clearQry = rs("rsQuery") = ""
%>
<script language="javascript">
var txtPredefined = '<%=getadminRepEditLngStr("LtxtPredefined")%>';
var txtCustomized = '<%=getadminRepEditLngStr("DtxtCustomized")%>';
function valFrm()
{
	if (document.form1.rsName.value == '')
	{
		alert('<%=getadminRepEditLngStr("LtxtValRepNam")%>');
		document.form1.rsName.focus();
		return false;
	}
	else if (document.form1.rsQuery.value == '')
	{
		alert('<%=getadminRepEditLngStr("LtxtValRepQry")%>');
		document.form1.rsQuery.focus();
		return false;
	}
	else if (document.form1.valRSQuery.value == 'Y')
	{
		alert("<%=getadminRepEditLngStr("LtxtValRepQryVal")%>");
		document.form1.btnVerfy.focus();
		return false;
	}
	return true;
}

var btnAfterVerfy;
var hdAfterVerfy;
function VerfyQuery()
{
	btnAfterVerfy = document.form1.btnVerfy;
	hdAfterVerfy = document.form1.valRSQuery;
	document.frmVerfyQuery.type.value = 'editQuery';
	document.frmVerfyQuery.rsIndex.value = '<%=Request("rsIndex")%>';
	document.frmVerfyQuery.Query.value = document.form1.rsQuery.value;
	document.frmVerfyQuery.varDataType.value = '';
	document.frmVerfyQuery.varQueryField.value = '';
	document.frmVerfyQuery.rsTop.value = document.form1.chkRSTop.checked ? 'Y' : 'N';
	document.frmVerfyQuery.submit();
}

function VerfyQueryVerified()
{
	btnAfterVerfy.src='images/btnValidateDis.gif'
	btnAfterVerfy.style.cursor = '';
	hdAfterVerfy.value='N';
}
function changeRepCmd(cmd)
{
	switch(cmd)
	{
	case 'variables':
		window.location.href='adminRepEdit.asp?rsIndex=<%=rsIndex%>&repCmd=variables&#tblRepVars';
		break;
	case 'repColor':
		<% If Not clearQry Then %>
		window.location.href='adminRepEdit.asp?rsIndex=<%=rsIndex%>&repCmd=repColor&#tblRepColors';
		<% Else %>
		alert('<%=getadminRepEditLngStr("LtxtEnterQryValSec")%>');
		return;
		<% End If %>
		break;
	default:
		window.location.href='adminRepEdit.asp?rsIndex=<%=rsIndex%>&repCmd=&#tblRepTotals';
	}
}

var NumUDField;
var NumUDTimerID = 0;
function doNumUD(udFld, dir)
{
	NumUDField = udFld;
	if (dir == 'U')
	{
		NumUDTimerID = setTimeout("doNumUDUp()", 250);
	}
	else
	{
		NumUDTimerID = setTimeout("doNumUDDown()", 250);
	}
}

function doNumUDUp()
{
	if(parseFloat(NumUDField.value)<32767)NumUDField.value=parseFloat(NumUDField.value)+1;
	NumUDTimerID = setTimeout("doNumUDUp()", 250);
}

function doNumUDDown()
{
	if(parseFloat(NumUDField.value)>0)NumUDField.value=parseFloat(NumUDField.value)-1;
	NumUDTimerID = setTimeout("doNumUDDown()", 250);
}

function stopDoNumUD()
{
	clearTimeout(NumUDTimerID);
}
function showLinkedTables(td)
{
	myTbl = document.getElementById('tblLinkedRep');
	myTbl.style.display='';
	myTbl.style.top = GetTopPos(td)+td.offsetHeight;
	myTbl.style.left = GetLeftPos(td)<% If Session("rtl") = "" Then %>-(myTbl.offsetWidth-td.offsetWidth)<% End If %>;
}
function hideLinkedTables()
{
	document.getElementById('tblLinkedRep').style.display='none';
}
function doSaveAs()
{
	if (document.form1.valRSQuery.value == 'Y')
	{
		alert("<%=getadminRepEditLngStr("LtxtValRepQryVal")%>");
		document.form1.btnVerfy.focus();
		return;
	}
	rsName = document.form1.rsName.value;
	rgIndex = document.form1.rgIndex.value;
	UserType = document.form1.UserType.value;
	OpenWin = this.open('', "CtrlWindow", "toolbar=no,menubar=no,location=no,scrollbars=yes,resizable=no, width=400,height=77");
	doMyLink('adminRepEditSaveAs.asp', 'rsName='+rsName+'&rgIndex=' + rgIndex + '&UserType=' + UserType + '&pop=Y', 'CtrlWindow');
}
function saveCopy(newRSName, newRGIndex)
{
	document.form1.saveAs.value = 'Y';
	document.form1.saveAsName.value = newRSName;
	document.form1.saveAsRG.value = newRGIndex;
	document.form1.submit();
}
</script>
<script language="javascript" src="js_up_down.js"></script>
<style type="text/css">
.style1 {
	text-align: center;
}
.style2 {
	COLOR: #31659C;
	FONT-SIZE: 10px;
	FONT-FAMILY: VERDANA;
	TEXT-DECORATION: none;
	BACKGROUND-COLOR: #E1F3FD;
	font-weight: bold;
	text-align: center;
}
</style>
<% 
ShowLinkedTables = False
sql = "select T2.rgName, T1.rsName " & _
		"from OLKRSTotals T0 " & _
		"inner join OLKRS T1 on T1.rsIndex = T0.rsIndex " & _
		"inner join OLKRG T2 on T2.rgIndex = T1.rgIndex " & _
		"where T0.linkType = 'R' and T0.linkObject = " & Request("rsIndex") & " " & _
		"order by 1, 2"
set rd = conn.execute(sql)
If Not rd.Eof Then
ShowLinkedTables = True %>
<table cellpadding="2" cellspacing="0" border="0" id="tblLinkedRep" style="position: absolute; display: none; z-index: 1;">
	<% do while not rd.eof %>
	<tr>
		<td class="TblRepNrm"><%=rd("rgName")%></td>
		<td class="TblRepNrm">-</td>
		<td class="TblRepNrm"><%=rd("rsName")%></td>
	</tr>
	<% rd.movenext
	loop %>
</table>
<% End If %>
<table border="0" cellpadding="0" width="100%" id="table3">
<form name="frmVerfyQuery" action="verfyQuery.asp" method="post" target="iVerfyQuery">
	<input type="hidden" name="type" value="editQuery">
	<input type="hidden" name="rsIndex" value="<%=Request("rsIndex")%>">
	<input type="hidden" name="Query" value="">
	<input type="hidden" name="UserType" value="<%=rs("UserType")%>">
	<input type="hidden" name="baseIndex" value="">
	<input type="hidden" name="parent" value="Y">
	<input type="hidden" name="varDataType" value="">
	<input type="hidden" name="varQueryField" value="">
	<input type="hidden" name="rsTop" value="">
</form>
<form name="form1" action="repSubmit.asp" method="POST" onsubmit="javascript:return valFrm();">
	<tr>
		<td>
		<table cellpadding="0" cellspacing="2" border="0" width="100%">
			<tr>
				<td class="TblRepTlt"><%=getadminRepEditLngStr("LttlEditRep")%></td>
				<% If ShowLinkedTables Then %>
				<td align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>" width="120" class="TblRepTlt" id="showLinkedRep" onmouseover="showLinkedTables(this);" onmouseout="hideLinkedTables();"><%=getadminRepEditLngStr("LtxtLinkedReports")%></td><% End If %>
			</tr>
		</table></td>
	</tr>
	<tr class="TblRepNrm">
		<td>
		<p align="justify"><img src="images/lentes.gif"><%=getadminRepEditLngStr("LttlEditRepNote")%></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table7">
			<tr>
				<td style="width: 100px" class="TblRepTlt">
				<%=getadminRepEditLngStr("DtxtName")%></td>
				<td colspan="7" class="TblRepNrm">
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td><input type="text" <% If CInt(Request("rsIndex")) < 0 Then %>readonly style="border-color: #afafaf; color: #808080;"<% End If %> name="rsName" size="88" value="<%=Server.HTMLEncode(rs("rsName"))%>" onkeydown="return chkMax(event, this, 60);"></td>
						<td valign="bottom"><a href="javascript:doFldTrad('RS', 'rsIndex', '<%=Request("rsIndex")%>', 'alterRSName', 'T', null);"><img src="images/trad.gif" alt="<%=getadminRepEditLngStr("DtxtTranslate")%>" border="0"></a></td>
					</tr>
				</table>
			    </td>
			</tr>
			<tr>
				<td valign="middle" style="width: 100px" class="TblRepTlt">
				<%=getadminRepEditLngStr("DtxtGroup")%></td>
				<% set rd = server.createobject("ADODB.RecordSet")
				sql = "select rgIndex, rgName from OLKRG where UserType = '" & UserType & "' order by rgName asc"
				set rd = conn.execute(sql)%>
				<td colspan="7" class="TblRepNrm"><select name="rgIndex<% If rs("rgIndex") < 0 Then %>2<% End If %>" <% If rs("rgIndex") < 0 Then %>disabled<% End If %> size="1">
				<% do while not rd.eof %>
				<option <% If CInt(rs("rgIndex")) = CInt(rd("rgIndex")) Then %>selected<% End If %> value="<%=rd("rgIndex")%>"><%=myHTMLEncode(rd("rgName"))%></option>
				<% rd.movenext
				loop
				set rd = nothing %>
				</select><% If rs("rgIndex") < 0 Then %><input type="hidden" name="rgIndex" value="<%=rs("rgIndex")%>"><% End If %>&nbsp;</td>
			</tr>
			<tr>
				<td style="width: 100px" class="TblRepTlt">
				<%=getadminRepEditLngStr("LtxtTime")%></td>
				<td width="26%">
				<select size="1" name="cmbRefresh">
				<option value="0"><%=getadminRepEditLngStr("DtxtDisabled")%></option>
				<option <% If rs("Refresh") = 1 Then %>selected<% End If %> value="1"><%=getadminRepEditLngStr("DtxtMinute")%></option>
				<option <% If rs("Refresh") = 5 Then %>selected<% End If %> value="5">5 <%=getadminRepEditLngStr("DtxtMinutes")%></option>
				<option <% If rs("Refresh") = 10 Then %>selected<% End If %> value="10">10 <%=getadminRepEditLngStr("DtxtMinutes")%></option>
				<option <% If rs("Refresh") = 30 Then %>selected<% End If %> value="30">30 <%=getadminRepEditLngStr("DtxtMinutes")%></option>
				<option <% If rs("Refresh") = 60 Then %>selected<% End If %> value="60"><%=getadminRepEditLngStr("DtxtHour")%></option>
				</select></td>
				<td width="15%" class="TblRepNrm">
				<input type="checkbox" name="rsTop" value="Y" <% If rs("rsTop") = "Y" Then %>checked<% End If %> id="chkRSTop" class="OptionButton" onclick="javascript:document.form1.btnVerfy.disabled=false;"><label for="chkRSTop"><%=getadminRepEditLngStr("DtxtTop")%> <span dir="ltr">(@top)</span></label></td>
				<td width="8%" class="TblRepNrm">
				<%=getadminRepEditLngStr("DtxtDefault")%></td>
				<td width="7%" class="TblRepNrm">
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td>
						<input id="rsTopDef" name="rsTopDef" value="<%=rs("rsTopDef")%>" size="6" onchange="if(!IsNumeric(this.value) || this.value=='')this.value=<%=rs("rsTopDef")%>" onfocus="this.select();" onmouseup="event.preventDefault()" maxlength="6"></td>
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
				<script language="javascript">NumUDAttachMin('form1', 'rsTopDef', 'btnNewOrderUp', 'btnNewOrderDown', 1);</script></td>
				<td width="15%" class="TblRepNrm">
				<input type="checkbox" name="LinkOnly" value="Y" <% If rs("LinkOnly") = "Y" Then %>checked<% End If %> id="chkLinkOnly" class="OptionButton"><label for="chkLinkOnly"><%=getadminRepEditLngStr("LtxtLinkOnly")%></label></td>
				<td width="8%" class="TblRepNrm">
				
				<input class="OptionButton" type="checkbox" name="chkActive" value="Y" <% If rs("Active") = "Y" Then %>checked<% End If %> id="chkActive"><label for="chkActive"><%=getadminRepEditLngStr("DtxtActive")%></label></td>
				<td class="TblRepNrm">
				
				&nbsp;</td>
			</tr>
			<tr>
				<td valign="top" style="width: 100px" class="TblRepTlt">
				<%=getadminRepEditLngStr("DtxtDescription")%></td>
				<td colspan="7" class="TblRepNrm">
				<table cellpadding="0" cellspacing="0" border="0" width="100%">
					<tr>
						<td><textarea rows="3" name="rsDesc" cols="86" style="width: 100%"><% If Not IsNull(rs("rsDesc")) Then %><%=myHTMLEncode(rs("rsDesc"))%><% End If %></textarea></td>
						<td valign="bottom" width="24"><a href="javascript:doFldTrad('RS', 'rsIndex', '<%=Request("rsIndex")%>', 'alterRSDesc', 'M', null);"><img src="images/trad.gif" alt="<%=getadminRepEditLngStr("DtxtTranslate")%>" border="0"></a></td>
					</tr>
				</table>
   				</td>
			</tr>
			<tr>
				<td valign="top" style="width: 100px" class="TblRepTlt">
				<%=getadminRepEditLngStr("DtxtQuery")%></td>
				<td colspan="7" class="TblRepNrm">
				<table cellpadding="0" cellspacing="0" border="0" width="100%">
					<tr>
						<td rowspan="2">
							<textarea dir="ltr" rows="16" name="rsQuery" cols="86" style="width: 100%" onkeydown="javascript:document.form1.btnVerfy.src='images/btnValidate.gif';document.form1.btnVerfy.style.cursor = 'hand';document.form1.valRSQuery.value='Y';"><%=myHTMLEncode(rs("rsQuery"))%></textarea>
							<div class="ui-widget" style="display: none;" id="txtQryErr">
								<div class="ui-state-error ui-corner-all"> 
									<p><span class="ui-icon ui-icon-alert" style="float: left;"></span> 
									<strong><%=getadminRepEditLngStr("DtxtQryErr")%>:</strong> <span id="txtQryErrMsg"></span></p>
								</div>
							</div>
						</td>
						<td valign="top" width="1">
							<img style="cursor: hand; " src="images/qry_note.gif" id="btnNoteFilter" alt="<%=getadminRepEditLngStr("DtxtDefinition")%>" onclick="javascript:doFldNote(18, 'rsQuery', '<%=Request("rsIndex")%>', null);">
						</td>
					</tr>
					<tr>
						<td valign="bottom" width="1">
							<img src="images/btnValidateDis.gif" id="btnVerfy" alt="<%=getadminRepEditLngStr("DtxtValidate")%>" onclick="javascript:if (document.form1.valRSQuery.value == 'Y')VerfyQuery();">
							<input type="hidden" name="valRSQuery" value="N">
						</td>
					</tr>
				</table>
    			</td>
			</tr>
			<tr>
				<td valign="top" style="width: 100px" class="TblRepTlt">
				<%=getadminRepEditLngStr("DtxtVariables")%></td>
				<td colspan="7" class="TblRepNrm">
				<% Select Case rs("UserType") 
					Case "C" %>
				<span dir="ltr">@CardCode</span> = <%=getadminRepEditLngStr("LtxtCCodeDesc")%>
				<% 	Case "V" %>
				<span dir="ltr">@SlpCode</span> = <%=getadminRepEditLngStr("LtxtACodeDesc")%>
				<% End Select %><br>
				<span dir="ltr">@LanID</span> = <%=getadminRepEditLngStr("DtxtLanID")%></td>
			</tr>
			<tr>
				<td valign="top" style="width: 100px" class="TblRepTlt">
				<%=getadminRepEditLngStr("DtxtFunctions")%></td>
				<td colspan="7" class="TblRepNrm"><% HideFunctionTitle = True
				functionClass = "TblRepNrm" %>
				<!--#include file="myFunctions.asp"--></td>
			</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table5">
			<tr>
				<td width="77">
				<input class="BtnRep" type="submit" value="<%=getadminRepEditLngStr("DtxtApply")%>" name="btnApply"></td>
				<td width="77">
				<input class="BtnRep" type="submit" value="<%=getadminRepEditLngStr("DtxtSave")%>" name="btnSave"></td>
				<td width="100">
				<input class="BtnRep" type="button" value="<%=getadminRepEditLngStr("DtxtSaveAs")%>" name="btnSaveAs" onclick="javascript:doSaveAs();" style="width: 100px; "></td>
				<td><hr size="1"></td>
				<% If CInt(Request("rsIndex")) < 0 Then %>
				<td width="77">
				<input type="button" class="BtnRep" value="<%=getadminRepEditLngStr("DtxtRestore")%>" name="btnRestore" onclick="javascript:if(confirm('<%=getadminRepEditLngStr("LtxtConfRestore")%>'))window.location.href='repSubmit.asp?uType=<%=UserType%>&cmd=repRestore&rsIndex=<%=Request("rsIndex")%>'"></td><% End If %>
				<td width="77">
				<input type="button" class="BtnRep" value="<%=getadminRepEditLngStr("DtxtCancel")%>" name="btnCancel" onclick="javascript:if(confirm('<%=getadminRepEditLngStr("DtxtConfCancel")%>'))window.location.href='adminReps.asp?uType=<%=UserType%>'"></td>
			</tr>
		</table>
		</td>
	</tr>
	<input type="hidden" name="repCmd" value="<%=Request("repCmd")%>">
	<input type="hidden" name="cmd" value="uRep">
	<input type="hidden" name="rsIndex" value="<%=rs("rsIndex")%>">
	<iframe id="iVerfyQuery" name="iVerfyQuery" style="display: none" src=""></iframe>
	<input type="hidden" name="UserType" value="<%=rs("UserType")%>">
	<input type="hidden" name="saveAs" value="N">
	<input type="hidden" name="saveAsName" value="">
	<input type="hidden" name="saveAsRG" value="">
	</form>
		<tr class="TblRepNrm">
		<td><select name="repCmd" size="1" onchange="javascript:changeRepCmd(this.value);">
		<option value=""><%=getadminRepEditLngStr("LoptFrmLnkTtl")%></option>
		<option value="variables" <% If Request("repCmd") = "variables" Then %>selected<% End If %>><%=getadminRepEditLngStr("DtxtVariables")%></option>
		<option value="repColor" <% If Request("repCmd") = "repColor" Then %>selected<% End If %>><%=getadminRepEditLngStr("DtxtColor")%></option></select></td>
	</tr>
	<% If Request("repCmd") = "variables" Then
		sql = 	"select varIndex, varName, varVar, " & _
 		"Case varType When 'T' Then N'" & getadminRepEditLngStr("LtxtText") & "' When 'Q' Then N'" & getadminRepEditLngStr("DtxtQuery") & "' When 'DD' Then '" & getadminRepEditLngStr("LtxtCmb") & "' When 'DP' Then N'" & getadminRepEditLngStr("DtxtDate") & "' When 'L' Then N'" & getadminRepEditLngStr("LtxtList") & "' When 'CL' Then N'" & getadminRepEditLngStr("LtxtChkList") & "' End varType, " & _
 		"varDataType, Ordr from OLKRSvars where rsIndex = " & Request("rsIndex") & " order by Ordr asc"
		rs.close
		rs.open sql, conn, 3, 1
		If Not rs.eof then %>
	<tr class="TblRepNrm">
		<td>
		<p align="justify"><img src="images/lentes.gif"> <%=getadminRepEditLngStr("LttlVarsNote")%></td>
	</tr>
	<script language="javascript">
	function valFrmAdmVars()
	{
		var varName = document.frmAdmVars.varName;
		if (varName)
		{
			if (varName.length)
			{
				for (var i = 0;i<varName.length;i++)
				{
					if (varName[i].value == '')
					{
						alert('<%=getadminRepEditLngStr("LtxtValVarNam")%>');
						varName[i].focus();
						return false;
					}
				}
			}
			else
			{
				if (varName.value == '')
				{
					alert('<%=getadminRepEditLngStr("LtxtValVarNam")%>');
					varName.focus();
					return false;
				}
			}
		}
		return true;
	}
	</script>
	<form name="frmAdmVars" action="repSubmit.asp" method="post" onsubmit="return valFrmAdmVars();">
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="tblRepVars">
			<tr class="TblRepTltSub">
				<td width="14"></td>
				<td width="200" class="style1"><%=getadminRepEditLngStr("DtxtName")%></td>
				<td width="120" class="style1"><%=getadminRepEditLngStr("DtxtVariable")%></td>
				<td width="120" class="style1"><%=getadminRepEditLngStr("LtxtFormat")%></td>
				<td width="120" class="style1"><%=getadminRepEditLngStr("DtxtType")%></td>
				<td class="style1" style="width: 80px"><%=getadminRepEditLngStr("DtxtOrder")%></td>
				<td></td>
			</tr>
			<% do while not rs.eof %>
			<input type="hidden" name="varIndex" value="<%=rs("varIndex")%>">
			<tr class="TblRepTbl">
				<td width="14">
				<b>
				<a href="adminRepEdit.asp?rsIndex=<%=Request("rsIndex")%>&editIndex=<%=rs("varIndex")%>&repCmd=variables&#editVar"><img border="0" src="images/<%=Session("rtl")%>flechaselec.gif" width="15" height="13"></a>
				</b>
				</td>
				<td width="200">
				<table cellpadding="0" cellspacing="0" border="0" width="100%">
					<tr class="TblRepTbl">
						<td><input type="text" name="varName<%=rs("varIndex")%>" id="varName" size="20" style="width: 100%;" value="<%=Server.HTMLEncode(rs("varName"))%>" onkeydown="return chkMax(event, this, 50);"></td>
						<td width="16"><a href="javascript:doFldTrad('RSVars', 'rsIndex,varIndex', '<%=Request("rsIndex")%>,<%=rs("varIndex")%>', 'alterVarName', 'T', null);"><img src="images/trad.gif" alt="<%=getadminRepEditLngStr("DtxtTranslate")%>" border="0"></a></td>
					</tr>
				</table>
				</td>
				<td width="120"><span dir="ltr">@<%=rs("varVar")%>&nbsp;</span></td>
				<td width="120"><%=rs("varType")%>&nbsp;</td>
				<td width="120"><%=rs("varDataType")%>&nbsp;</td>
				<td align="center" style="width: 80px"><table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td>
						<input id="Ordr<%=rs("varIndex")%>" name="Ordr<%=rs("varIndex")%>" value="<%=rs("Ordr")%>" size="6" onchange="if(!IsNumeric(this.value) || this.value=='')this.value=<%=rs("Ordr")%>" onfocus="this.select();" onmouseup="event.preventDefault()" maxlength="6"></td>
						<td valign="middle">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td><img src="images/img_nud_up.gif" id="Ordr<%=rs("varIndex")%>Up"></td>
							</tr>
							<tr>
								<td><img src="images/spacer.gif"></td>
							</tr>
							<tr>
								<td><img src="images/img_nud_down.gif" id="Ordr<%=rs("varIndex")%>Down"></td>
							</tr>
						</table>
						</td>
					</tr>
				</table>
				<script language="javascript">NumUDAttachMin('frmAdmColors', 'Ordr<%=rs("varIndex")%>', 'Ordr<%=rs("varIndex")%>Up', 'Ordr<%=rs("varIndex")%>Down', 1);</script></td>
				<td>
				<a href="javascript:if(confirm('<%=getadminRepEditLngStr("LtxtConfDelVar")%>'.replace('{0}', '<%=Replace(rs("varName"), "'", "\'")%>')))window.location.href='repSubmit.asp?cmd=remRSVar&rsIndex=<%=Request("rsIndex")%>&varIndex=<%=rs("varIndex")%>'">
			<img border="0" src="images/remove.gif"></a></td>
			</tr>
			<% rs.movenext
			loop %>
			</table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table9">
			<tr>
				<td width="77">
				<input class="BtnRep" type="submit" value="<%=getadminRepEditLngStr("DtxtSave")%>" name="btnSave"></td>
				<td><hr size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
	<input type="hidden" name="cmd" value="admVars">
	<input type="hidden" name="rsIndex" value="<%=rsIndex%>">
	<input type="hidden" name="UserType" value="<%=UserType%>">
</form>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table5">
			<tr>
				<td></td>
			</tr>
		</table>
		</td>
	</tr>
	<% end if 
If Request("editIndex") <> "" Then
	sql = "select *, " & _
	"Case When Exists(select 'A' from OLKRSVarsBase where rsIndex = T0.rsIndex and varIndex = T0.varIndex) Then 'Y' Else 'N' End IsTarget " & _
	"from OLKRSVars T0 where rsIndex = " & Request("rsIndex") & " and varIndex = " & Request("editIndex")
	set rs = conn.execute(sql)
	varName = rs("varName")
	varVar = rs("varVar")
	vType = rs("varType")
	varDataType = rs("varDataType")
	varMaxChar = rs("varMaxChar")
	varDefVars = rs("varDefVars")
	varNotNull = rs("varNotNull")
	varShowRep = rs("varShowRep")
	varDefBy = rs("DefValBy")
	varDefValue = rs("DefValValue")
	varDefDate = rs("DefValDate")
	Ordr = rs("Ordr")
	IsTarget = rs("IsTarget") = "Y"
	Select Case varDefVars
		Case "Q"
			varQuery = rs("varQuery")
			varQueryField = rs("varQueryField")
		Case "F"
			sql = "select valValue + + ',' + + valText As 'Line' from OLKRSVarsVals where rsIndex = " & Request("rsIndex") & " and varIndex = " & Request("editIndex")
			rs.close
			rs.open sql, conn, 3, 1
			do while not rs.eof
			if rs.bookmark > 1 then varQuery = varQuery & VbNewLine
			varQuery = varQuery & rs("Line")
			rs.movenext
			loop
			rs.close
	End Select
Else
	sql = "select IsNull((select Max(Ordr)+1 from OLKRSVars where rsIndex = " & Request("rsIndex") & "), 0) Ordr"
	set rs = conn.execute(sql)
	Ordr = rs("Ordr")
	varName = ""
	varDefBy = "N"
	IsTarget = False
End IF
Function CleanItem(Value)
	CleanItem = Replace(Replace(Replace(myHTMLEncode(Value),"#","%23"),"&","%26"),"""","%22")
End Function
	%>
<form name="form2" action="repSubmit.asp" method="POST" onsubmit="return valFrmVar();">
<% If Request("editIndex") = "" Then %>
<input type="hidden" name="varNameTrad">
<input type="hidden" name="varQueryDef">
<input type="hidden" name="varDefValueDef">
<% End If %>
	<tr class="TblRepTlt" id="editVar">
		<td>&nbsp;<% If Request("editIndex") = "" Then %><%=getadminRepEditLngStr("LttlAddVar")%><% Else %><%=getadminRepEditLngStr("LttlEditVar")%><% End If %></td>
	</tr>
	<tr class="TblRepNrm">
		<td>
		<p align="justify"><img src="images/lentes.gif"><%=getadminRepEditLngStr("LttlVarNote")%></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table8">
			<tr>
				<td width="97" class="TblRepTlt">
					<%=getadminRepEditLngStr("DtxtName")%></td>
				<td width="224" class="TblRepNrm">
				
    			<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td><input type="text" name="varName" size="20" value="<%=Server.HTMLEncode(varName)%>" onkeydown="return chkMax(event, this, 50);">
						</td>
						<td width="16"><a href="javascript:doFldTrad('RSVars', 'rsIndex,varIndex', '<%=Request("rsIndex")%>,<%=Request("editIndex")%>', 'alterVarName', 'T', <% If Request("editIndex") = "" Then %>null<% Else %>document.form2.varNameTrad<% End If %>);"><img src="images/trad.gif" alt="<%=getadminRepEditLngStr("DtxtTranslate")%>" border="0"></a></td>
					</tr>
				</table>
    			
    			</td>
				<td width="131" class="TblRepTlt">
				<%=getadminRepEditLngStr("LtxtFormat")%></td>
				<td class="TblRepNrm">
			    <select size="1" name="varType" onchange="changeType(this.value)">
				<option <% If vType = "T" Then %>selected<% End If %> value="T">
				<%=getadminRepEditLngStr("LtxtText")%></option>
				<option <% If vType = "Q" Then %>selected<% End If %> value="Q">
				<%=getadminRepEditLngStr("DtxtQuery")%></option>
				<option <% If vType = "DD" Then %>selected<% End If %> value="DD">
				<%=getadminRepEditLngStr("LtxtCmb")%></option>
				<option <% If vType = "DP" Then %>selected<% End If %> value="DP">
				<%=getadminRepEditLngStr("DtxtDate")%></option>
				<option <% If vType = "L" Then %>selected<% End If %> value="L">
				<%=getadminRepEditLngStr("LtxtList")%></option>
				<option <% If vType = "CL" Then %>selected<% End If %> value="CL">
				<%=getadminRepEditLngStr("LtxtChkList")%>
				</option>
				</select></td>
			</tr>
			<tr>
				<td width="97" class="TblRepTlt"><span dir="ltr">@<%=getadminRepEditLngStr("DtxtVariable")%></span></td>
				<td width="224" class="TblRepNrm">
			    <input type="text" name="varVar" size="20" value="<%=varVar%>" maxlength="50" onkeydown="javascript:document.form2.btnVerfyVarVar.src='images/btnValidate.gif';document.form2.btnVerfyVarVar.style.cursor = 'hand';document.form2.valVarVar.value='Y';">
				<img src="images/btnValidateDis.gif" id="btnVerfyVarVar" alt="<%=getadminRepEditLngStr("DtxtValidate")%>" onclick="javascript:if (document.form2.valVarVar.value == 'Y')VerfyVarVar();">
				<input type="hidden" name="valVarVar" value="N">
				</td>
				<td width="131" class="TblRepTlt">
				<%=getadminRepEditLngStr("DtxtType")%></td>
				<td class="TblRepNrm">
				
    <select size="1" name="varDataType" onchange="chkTypeDate(this)">
	<option <% If varDataType = "nvarchar" Then %>selected<% End If %> value="nvarchar"><%=getadminRepEditLngStr("LtxtText")%></option>
	<option <% If varDataType = "datetime" Then %>selected<% End If %> value="datetime"><%=getadminRepEditLngStr("DtxtDate")%></option>
	<option <% If varDataType = "float" or varDataType = "numeric" Then %>selected<% End If %> value="numeric"><%=getadminRepEditLngStr("DtxtNumeric")%></option>
	<option <% If varDataType = "int" Then %>selected<% End If %> value="int"><%=getadminRepEditLngStr("LtxtNumWhole")%></option>
	</select></td>
			</tr>
			<tr>
				<td class="TblRepNrm">
				&nbsp;
				</td>
				<td class="TblRepNrm">
				<table border="0" cellpadding="0" cellspacing="0">
					<tr>
						<td>
						<input type="checkbox" name="varNotNull" value="Y" <% If varNotNull = "Y" or varNotNull = "" Then %>checked<% End If %> id="varNotNull" class="OptionButton" style="background:background-image"></td>
						<td class="TblRepNrm"><label for="varNotNull"><%=getadminRepEditLngStr("DtxtNotNull")%></label></td>
					</tr>
				</table>
				</td>
				<td class="TblRepTlt" width="131">
				<%=getadminRepEditLngStr("LtxtMaxChar")%></td>
				<td class="TblRepNrm">
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td>
						<input id="varMaxChar" name="varMaxChar" value="<%=varMaxChar%>" size="6" onfocus="this.select();" onmouseup="event.preventDefault()" maxlength="6"></td>
						<td valign="middle">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td><img src="images/img_nud_up.gif" id="btnVarMaxCharUp"></td>
							</tr>
							<tr>
								<td><img src="images/spacer.gif"></td>
							</tr>
							<tr>
								<td><img src="images/img_nud_down.gif" id="btnVarMaxCharDown"></td>
							</tr>
						</table>
						</td>
					</tr>
				</table>
			    </td>
			</tr>
			<tr class="TblRepNrm">
				<td>&nbsp;</td>
				<td>
				<table border="0" cellpadding="0" cellspacing="0">
					<tr>
						<td>
						<input style="background:background-image" class="OptionButton" type="checkbox" name="varShowRep" value="Y" <% If varShowRep = "Y" or varShowRep = "" Then %>checked<% End If %> id="varShowRep"></td>
						<td class="TblRepNrm"><label for="varShowReP"><%=getadminRepEditLngStr("LtxtShowInRep")%></label></td>
					</tr>
				</table>
				</td>
				<td class="TblRepTlt" width="131">
				<%=getadminRepEditLngStr("DtxtOrder")%></td>
				<td><table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td>
						<input id="Ordr" name="Ordr" value="<%=Ordr%>" size="6" onchange="if(!IsNumeric(this.value) || this.value=='')this.value=<%=Ordr%>" onfocus="this.select();" onmouseup="event.preventDefault()" maxlength="6"></td>
						<td valign="middle">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td><img src="images/img_nud_up.gif" id="OrdrUp"></td>
							</tr>
							<tr>
								<td><img src="images/spacer.gif"></td>
							</tr>
							<tr>
								<td><img src="images/img_nud_down.gif" id="OrdrDown"></td>
							</tr>
						</table>
						</td>
					</tr>
				</table>
				<script language="javascript">NumUDAttachMin('form2', 'Ordr', 'OrdrUp', 'OrdrDown', 1);</script></td>
			</tr>
			<tr>
				<td valign="top" colspan="4">
				<table border="0" width="100%" id="table12" cellspacing="2" cellpadding="0">
					<tr>
						<td class="TblRepNrm">
				<input <% If varDefVars = "Q" or varDefVars = "" Then %>checked<% End If %> type="radio" value="Q" name="varQueryBy" id="fp1" class="OptionButton" style="background:background-image" onclick="doChangeDefVars('Q');"><label for="fp1"><%=getadminRepEditLngStr("DtxtQuery")%></label><input <% If varDefVars = "F" Then %>checked<% End If %> type="radio" name="varQueryBy" value="F" id="fp2" class="OptionButton" style="background:background-image" onclick="doChangeDefVars('F');"><label for="fp2"><%=getadminRepEditLngStr("LtxtFixVals")%> </label>(<%=getadminRepEditLngStr("LtxtValText")%>)</td>
						<td width="200" class="TblRepTlt"><%=getadminRepEditLngStr("LtxtBaseVars")%></td>
						<td class="TblRepTlt"><%=getadminRepEditLngStr("LtxtSelFld")%></td>
					</tr>
					<tr>
						<td width="440" class="TblRepNrm">
						<table cellpadding="0" cellspacing="0" border="0" width="100%">
							<tr>
								<td rowspan="2">
									<textarea dir="ltr" rows="8" name="varQuery" id="varQuery" cols="83" class="input" <% If (vType <> "DD" and vType <> "L" and vType <> "Q" and vType <> "CL") or vType = "" Then %>disabled style="background-color: #CCCCCC"<% End If %> onkeydown="javascript:document.form2.btnVerfyVar.src='images/btnValidate.gif';document.form2.btnVerfyVar.style.cursor = 'hand';document.form2.valVarQuery.value='Y';"><%=myHTMLEncode(varQuery)%></textarea>
								</td>
								<td valign="top" width="1">
									<img style="cursor: hand; " src="images/qry_note.gif" id="btnNoteFilter" alt="<%=getadminRepEditLngStr("DtxtDefinition")%>" onclick="javascript:doFldNote(18, 'varQuery', '<%=Request("rsIndex")%><%=Request("editIndex")%>', <% If Request("editIndex") <> "" Then %>null<% Else %>document.form2.varQueryDef<% End If %>);">
								</td>
							</tr>
							<tr>
								<td valign="bottom" width="1">
									<img src="images/btnValidateDis.gif" id="btnVerfyVar" alt="<%=getadminRepEditLngStr("DtxtValidate")%>" onclick="javascript:if (document.form2.valVarQuery.value == 'Y')VerfyVarQuery();">
									<input type="hidden" name="valVarQuery" value="N">
								</td>
							</tr>
						</table>
    					</td>
						<td valign="top" width="200" class="TblRepNrm">
						<%
						sql = "select T0.varIndex, '@' + T0.varVar VarVar, T0.varName VarDesc, "
						
						If Request("editIndex") = "" Then 
							sql = sql & " 'N' Verfy "
						Else
							sql = sql & " Case When Exists(select 'A' from OLKRSVarsBase where rsIndex = " & Request("rsIndex") & " and varIndex = " & Request("editIndex") & " and baseIndex = T0.varIndex) Then 'Y' Else 'N' End Verfy, " & _
							"varDataType, varMaxChar "
						End If
						
						sql = sql & " " & _
						"from OLKRSVars T0 " & _
						"where rsIndex = " & Request("rsIndex")
				If Request("editIndex") <> "" Then sql = sql & " and varIndex <> " & Request("editIndex")
					set rs = conn.execute(sql)
					if not rs.eof then
					enabledBaseVars = True %>
				<table cellpadding="0" cellspacing="0" border="0" style="border: 1px solid #4783C5;">
				<% do while not rs.eof %>
					<tr class="TblRepNrm">
						<td>
							<input style="background:background-image" onclick="javascript:chkBase(this);" <% If Request("editIndex") = "" or Request("editIndex") <> "" and ((vType <> "Q" and vType <> "L" and vType <> "DD") or vType = "") and (varDefVars <> "Q" or varDefVars <> "") Then %>disabled<% End If %> class="OptionButton" type="checkbox" name="baseVar" id="baseVars<%=rs("varIndex")%>" value="<%=rs("varIndex")%>" <% If rs("Verfy") = "Y" Then %>checked<% End If %> onclick="javascript:document.form2.btnVerfyVar.disabled=false;"><label for="baseVars<%=rs("varIndex")%>"><span dir="ltr"><%=rs("VarVar")%></span> 
							- <%=rs("VarDesc")%></label>&nbsp;&nbsp;
							<input type="hidden" name="baseIndex" value="<%=rs("varIndex")%>">
						</td>
					</tr>
					<% rs.movenext
					loop %>
				</table>
				<% Else
				enabledBaseVars = False
				End If %>
						&nbsp;</td>
						<td valign="top" class="TblRepNrm">
						<select size="1" name="varQueryField" <% If ((vType <> "Q") or vType = "") and (varDefVars <> "Q" or varDefVars <> "") Then %>disabled<% End If %> onchange="javascript:document.form2.btnVerfyVar.disabled=false;">
						<% 
						If varQuery <> "" and varDefVars = "Q" Then
							set rTest = Server.CreateObject("ADODB.RecordSet")
							testQuery = "declare @LanID int set @LanID = " & Session("LanID") & " "
							Select Case UserType 
								Case "C" 
									testQuery = testQuery & "declare @CardCode nvarchar(15) "
								Case "V" 
								testQuery = testQuery & "declare @SlpCode int "
							End Select
							If enabledBaseVars Then
								rs.movefirst
								do while not rs.eof
									If rs("varDataType") = "nvarchar" Then 
										MaxVar = "(" & rs("varMaxChar") & ")"
									ElseIf rs("varDataType") = "numeric" Then
										MaxVar = "(19,6)"
									Else
										MaxVar = ""
									End If
									testQuery = testQuery & "declare " & rs("varVar") & " " & rs("varDataType") & " " & MaxChar & " "
								rs.movenext
								loop
							End If
							testQuery = testQuery & varQuery
							On Error Resume Next
							set rTest = conn.execute(QueryFunctions(testQuery))
							If Err.Number = 0 Then
								For each itm in rTest.Fields
								If itm.Name <> "" Then %>
								<option <% If varQueryField = itm.Name Then %>selected<% End If %> value="<%=myHTMLEncode(itm.Name)%>"><%=myHTMLEncode(itm.Name)%></option>
								<% End If
								Next
								set rTest = Nothing
							Else
						%><script type="text/javascript">
						$(document).ready(function() {
						document.getElementById('txtQryVarErr').style.display = '';
						document.getElementById('txtQryVarErrMsg').innerText = '<%=Replace(Err.Description, "'", "\'")%>';
					});</script><%
							End If
						End If %>
						</select></td>
					</tr>
					<tr>
						<td colspan="3" style="display: none;" id="txtQryVarErr">
						<div class="ui-widget">
							<div class="ui-state-error ui-corner-all"> 
								<p><span class="ui-icon ui-icon-alert" style="float: left;"></span> 
								<strong><%=getadminRepEditLngStr("DtxtQryErr")%>:</strong> <span id="txtQryVarErrMsg"></span></p>
							</div>
						</div>
						</td>
					</tr>
					<tr>
						<td>
						<table cellpadding="0" cellspacing="2" border="0" width="100%">
							<tr>
								<td class="TblRepTlt"><%=getadminRepEditLngStr("LtxtDefValue")%></td>
								<td class="TblRepNrm">
								<input type="radio" name="varDefBy" id="varDefByN" value="N" <% If varDefBy = "N" Then %>checked<% End If %> class="OptionButton" style="background:background-image" onclick="javascript:changeDefVarBy(this.value);"><label for="varDefByN"><%=getadminRepEditLngStr("LtxtNone")%></label>
								<input <% If IsTarget Then %>disabled<% End If %> type="radio" name="varDefBy" id="varDefByV" value="V" <% If varDefBy = "V" Then %>checked<% End If %> class="OptionButton" style="background:background-image" onclick="javascript:changeDefVarBy(this.value);"><label for="varDefByV"><%=getadminRepEditLngStr("DtxtValue")%></label>
								<input <% If IsTarget Then %>disabled<% End If %> type="radio" name="varDefBy" id="varDefByQ" value="Q" <% If varDefBy = "Q" Then %>checked<% End If %> class="OptionButton" style="background:background-image" onclick="javascript:changeDefVarBy(this.value);"><label for="varDefByQ"><%=getadminRepEditLngStr("DtxtQuery")%></label>
								</td>
							</tr>
						</table>
						</td>
						<td valign="top" width="200" class="TblRepNrm">
						&nbsp;</td>
						<td valign="top" class="TblRepNrm">
						&nbsp;</td>
					</tr>
					<tr class="TblRepNrm">
						<td width="440">
						<table cellpadding="0" cellpadding="0" border="0" id="tblDefValValue" style="<% If varDefBy <> "V" or varDefBy = "V" and vType = "DP" Then %>display:none;<% End If %>">
							<tr>
								<td><input type="text" name="varDefValValue" id="varDefValValue" size="20" value="<% If varDefBy = "V" and varDataType <> "datetime" and not IsNull(varDefValue) Then %><%=Server.HTMLEncode(varDefValue)%><% End If %>"></td>
							</tr>
						</table>
						<table cellpadding="0" cellpadding="0" border="0" id="tblDefValDate" style="<% If varDefBy <> "V"  or varDefBy = "V" and vType <> "DP" Then %>display:none;<% End If %>">
							<tr>
								<td><img border="0" src="images/cal.gif" id="btnDefValDateImg" width="16" height="16" style="float:left;padding-left:1px;padding-top:1px"></td>
								<td><input type="text" name="varDefValDate" id="varDefValDate" readonly size="20" onclick="btnDefValDateImg.click();" value="<% If varDefBy = "V" and varDataType = "datetime" Then %><%=FormatDate(varDefDate, False)%><% End If %>"></td>
							</tr>
						</table>
						<table cellpadding="0" cellspacing="0" border="0" width="100%" id="tblDefValQuery" style="<% If varDefBy <> "Q" Then %>display:none;<% End If %>">
							<tr>
								<td rowspan="2">
									<textarea dir="ltr" rows="4" name="varDefValQuery" cols="83" onkeydown="javascript:document.form2.btnDefValQuery.src='images/btnValidate.gif';document.form2.btnDefValQuery.style.cursor = 'hand';document.form2.valDefValQuery.value='Y';"><% If varDefBy = "Q" Then %><%=myHTMLEncode(varDefValue)%><% End If %></textarea>
								</td>
								<td valign="top" width="1">
									<img style="cursor: hand; " src="images/qry_note.gif" id="btnNoteFilter" alt="<%=getadminRepEditLngStr("DtxtDefinition")%>" onclick="javascript:doFldNote(18, 'varDefValue', '<%=Request("rsIndex")%><%=Request("editIndex")%>', <% If Request("editIndex") <> "" Then %>null<% Else %>document.form2.varDefValueDef<% End If %>);">
								</td>
							</tr>
							<tr>
								<td valign="bottom" width="1">
									<img src="images/btnValidateDis.gif" id="btnDefValQuery" alt="<%=getadminRepEditLngStr("DtxtValidate")%>" onclick="javascript:if (document.form2.valDefValQuery.value == 'Y')VerfyVarDefQuery();">
									<input type="hidden" name="valDefValQuery" value="N">
								</td>
							</tr>
						</table>
						</td>
						<td valign="top" width="200">
						&nbsp;</td>
						<td valign="top">
						&nbsp;</td>
					</tr>
				</table>
				</td>
			</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table9">
			<tr>
				<td width="77">
				<input class="BtnRep" type="submit" value="<%=getadminRepEditLngStr("DtxtApply")%>" name="btnApply"></td>
				<td width="77">
				<input class="BtnRep" type="submit" value="<%=getadminRepEditLngStr("DtxtSave")%>" name="btnSave"></td>
				<td><hr size="1"></td>
				<td width="77">
				<input type="button" class="BtnRep" value="<%=getadminRepEditLngStr("DtxtCancel")%>" name="btnCancel" onclick="javascript:if(confirm('<%=getadminRepEditLngStr("DtxtConfCancel")%>'))window.location.href='adminRepEdit.asp?rsIndex=<%=Request("rsIndex")%>&#editVar'"></td>
			</tr>
		</table>
		</td>
	</tr>
	<input type="hidden" name="cmd" value="<% If Request("EditIndex") = "" Then %>add<% Else %>edit<% End If %>Var">
<input type="hidden" name="rsIndex" value="<%=rsIndex%>">
<input type="hidden" name="varIndex" value="<%=Request("editIndex")%>">
</form>
<script language="javascript">
NumUDAttachMin('form2', 'varMaxChar', 'btnVarMaxCharUp', 'btnVarMaxCharDown', 1);

Calendar.setup({
    inputField     :    "varDefValDate",     // id of the input field
    ifFormat       :    "<%=GetCalendarFormatString%>",      // format of the input field
    button         :    "btnDefValDateImg",  // trigger for the calendar (button ID)
    align          :    "Tr",           // alignment (defaults to "Bl")
    singleClick    :    true
});
function chkBase(chk)
{
	if (getDefBy() != 'N')
	{
		if (confirm('<%=getadminRepEditLngStr("LtxtChkBaseVars")%>'))
		{
			document.form2.varDefByN.checked = true;
			changeDefVarBy(getDefBy());
			document.form2.varDefByV.disabled = true;
			document.form2.varDefByQ.disabled = true;
		}
		else
		{
			chk.checked = false;
		}
	}
	else
	{
		document.form2.varDefByV.disabled = getBaseIndex() != '';
		document.form2.varDefByQ.disabled = getBaseIndex() != '';
	}
}
function changeDefVarBy(val)
{
	varType = document.form2.varType.value;
	document.getElementById('tblDefValValue').style.display = val == 'V' && varType != 'DP' ? '' : 'none';
	document.getElementById('tblDefValDate').style.display = val == 'V' && varType == 'DP' ? '' : 'none';
	document.getElementById('tblDefValQuery').style.display = val == 'Q' ? '' : 'none';
}
function VerfyVarVar()
{
	btnAfterVerfy = document.form2.btnVerfyVarVar;
	hdAfterVerfy = document.form2.valVarVar;
	document.frmVerfyQuery.type.value = 'newQueryVar';
	document.frmVerfyQuery.rsIndex.value = '<%=Request("rsIndex")%>';
	document.frmVerfyQuery.Query.value = document.form2.varVar.value;
	document.frmVerfyQuery.baseIndex.value = '';
	document.frmVerfyQuery.submit();
}

function getBaseIndex()
{

	var retVal = '';
	<% If enabledBaseVars Then %>
		if (document.form2.baseIndex.length)
		{
			for (var i = 0;i<document.form2.baseIndex.length;i++)
			{
				if (document.getElementById('baseVars' + document.form2.baseIndex[i].value).checked)
				{
					if (retVal != '') retVal += ', ';
					retVal += document.getElementById('baseVars' + document.form2.baseIndex[i].value).value;
				}
			}
		}
		else
		{
			if (document.getElementById('baseVars' + document.form2.baseIndex.value).checked) retVal = document.getElementById('baseVars' + document.form2.baseIndex.value).value;
		}
	<% End If %>
	return retVal;
}

function VerfyVarDefQuery()
{
	btnAfterVerfy = document.form2.btnDefValQuery;
	hdAfterVerfy = document.form2.valDefValQuery;
	document.frmVerfyQuery.type.value = 'RSVarDef';
	document.frmVerfyQuery.rsIndex.value = '<%=Request("rsIndex")%>';
	document.frmVerfyQuery.Query.value = document.form2.varDefValQuery.value;
	document.frmVerfyQuery.baseIndex.value = getBaseIndex();
	document.frmVerfyQuery.varDataType.value = document.form2.varDataType.value;
	document.frmVerfyQuery.varQueryField.value = '';
	document.frmVerfyQuery.submit();
}
function VerfyVarQuery()
{
	varQueryBy = document.form2.varQueryBy;
	if (varQueryBy[0].checked)
	{
		btnAfterVerfy = document.form2.btnVerfyVar;
		hdAfterVerfy = document.form2.valVarQuery;
		document.frmVerfyQuery.type.value = 'RSVar';
		document.frmVerfyQuery.rsIndex.value = '<%=Request("rsIndex")%>';
		document.frmVerfyQuery.Query.value = document.form2.varQuery.value;
		document.frmVerfyQuery.baseIndex.value = getBaseIndex();
		document.frmVerfyQuery.varDataType.value = document.form2.varDataType.value;
		document.frmVerfyQuery.varQueryField.value = document.form2.varType.value == 'Q' ? document.form2.varQueryField.value : '';
		document.frmVerfyQuery.submit();
	}
	else
	{
		var myLines = document.form2.varQuery.value.split('\n');
		for (var i = 0;i<myLines.length;i++)
		{
			if (myLines[i].length == 1) {
				alert('<%=getadminRepEditLngStr("LtxtValLineNoData")%>'.replace('{0}', (i+1)));
				return; }
			else if (myLines[i].split(',').length != 2) {
				alert('<%=getadminRepEditLngStr("LtxtLineValsWrongData")%>'.replace('{0}', (i+1)));
				return; }
		}
		document.form2.btnVerfyVar.src='images/btnValidateDis.gif';
		document.form2.btnVerfyVar.style.cursor = '';
		document.form2.valVarQuery.value='N';
	}	
}
function getSqlQueryField()
{
	return document.form2.varQueryField;
}
function doChangeDefVars(DefVar)
{
	var dType = document.form2.varType.value;
	document.form2.varQueryField.disabled = !(dType == 'Q' && DefVar == 'Q');
	<% If enabledBaseVars Then %>
		if (document.form2.baseIndex.length)
		{
			for (var i = 0;i<document.form2.baseIndex.length;i++)
			{
				document.getElementById('baseVars' + document.form2.baseIndex[i].value).disabled = !((dType == 'Q' || dType == 'L' || dType == 'DD' || dType == 'CL') && DefVar == 'Q');
			}
		}
		else
		{
			document.getElementById('baseVars' + document.form2.baseIndex.value).disabled = !((dType == 'Q' || dType == 'L' || dType == 'DD' || dType == 'CL') && DefVar == 'Q');
		}
	<% End If %>
}
</script>
<% ElseIf Request("repCmd") = "" Then
On Error Resume Next
sql = getGenQry
set rs = conn.execute(sql)
If Err.Number = 0 Then
	sql = ""
	For i = 1 to rs.Fields.count
		If i > 1 Then sql = sql & " union "
		sql = sql & "select " & i & " ID, N'" & Replace(rs.Fields(i-1).Name, "'", "''") & "' colName, " & rs.Fields(i-1).Type & " colType "
	next
	If Not clearQry Then
		sql =	"select T0.ID, T0.colName, T0.colType, " & _
				"Case When colType in (16, 2, 3, 20, 17, 18, 19, 21, 4, 5, 14, 131, 139) Then 'True' Else 'False' End EnableFormat, " & _
				"colTotal, colAlign, colFormat, colSum, colNB, colShow, IsNull(linkType,'N') linkType, " & _
				"Case When Exists(select '' from OLKRSColors where rsIndex = " & Request("rsIndex") & " and (colName = T0.colName or colOpBy = 'F' and colValue = T0.colName)) Then 'Y' Else 'N' End ShowColor, " & _
				"Case When Exists(select '' from OLKRSLinksVars where rsIndex = " & Request("rsIndex") & " and valBy = 'F' and valValue = T0.colName) Then 'Y' Else 'N' End ShowLinkVar " & _
				"from (" & sql & ") T0 " & _
				"left outer join OLKRSTotals T1 on T1.rsIndex = " & Request("rsIndex") & " and T1.colName = T0.colName "
		set rs = conn.execute(sql)
	End If
Else
	%><script type="text/javascript">
	$(document).ready(function() {

	document.getElementById('txtQryErr').style.display = '';
	document.getElementById('txtQryErrMsg').innerText = '<%=Replace(Err.Description, "'", "\'")%>';
});</script><%
	clearQry = True
End If
%>
<form name="form3" action="repSubmit.asp" method="POST" onsubmit="return valFrm3();">
	<tr class="TblRepNrm">
		<td>
		<p align="justify"><img src="images/lentes.gif"><%=getadminRepEditLngStr("LttlFrmLnkTtlNote")%></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="tblRepTotals">
			<tr class="TblRepTltSub">
				<td class="style1"><%=getadminRepEditLngStr("DtxtField")%></td>
				<td class="style1"><%=getadminRepEditLngStr("DtxtType")%></td>
				<td class="style1"><%=getadminRepEditLngStr("DtxtAlignment")%></td>
				<td class="style1"><%=getadminRepEditLngStr("LtxtFormat")%></td>
				<td class="style1"><%=getadminRepEditLngStr("DtxtPosition2")%></td>
				<td class="style1"><%=getadminRepEditLngStr("LtxtTotal")%></td>
				<td class="style1"><nobr><%=getadminRepEditLngStr("LtxtSumCol")%></nobr></td>
				<td class="style1"><nobr><%=getadminRepEditLngStr("LtxtNoBR")%></nobr></td>
				<td class="style1"><%=getadminRepEditLngStr("DtxtLink")%></td>
				<td style="width: 16px"></td>
				<td style="width: 18px"></td>
			</tr>
			<% 
			If Not clearQry Then
			cCount = 0
			foundNumInSale = "false"
			foundNumInBuy = "false"
			do while not rs.eof
			cCount = cCount + 1
			colName = Replace(Replace(rs("colName")," ",""), ".", "")
			If LCase(colName) = "numinsale" and getColType(rs("ColType")) = "adNumeric" Then foundNumInSale = "true"
			If LCase(colName) = "numinbuy" and getColType(rs("ColType")) = "adNumeric" Then foundNumInBuy = "true" %>
			<input type="hidden" name="colName" value="'<% If Not IsNull(rs("colName")) Then %><%=Server.HTMLEncode(rs("colName"))%><% End If %>'">
			<input type="hidden" name="arrColName" value="N'<%=Replace(Server.HTMLEncode(rs("colName")), "'", "''")%>'">
			<tr class="TblRepTbl">
				<td valign="top">
				<table cellpadding="0" cellspacing="0" border="0" width="100%">
					<tr class="TblRepTbl">
						<td><%=rs("colName")%>&nbsp;</td>
						<td align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>" style="width: 16px">
						<a href="javascript:doFldTrad('RSTotals', 'rsIndex,colName', '<%=Request("rsIndex")%>,<%=colName%>', 'alterColName', 'T', null);"><img src="images/trad.gif" alt="<%=getadminRepEditLngStr("DtxtTranslate")%>" border="0"></a></td>
					</tr>
				</table>
				</td>
				<td width="100"><%=getColType(rs("ColType"))%>&nbsp;</td>
				<td width="90">
				<select size="1" name="Align<%=rs("ID")%>">
				<option></option>
				<option <% If rs("colAlign") = "L" Then %>selected<% End If %> value="L"><%=getadminRepEditLngStr("DtxtLeft")%></option>
				<option <% If rs("colAlign") = "C" Then %>selected<% End If %> value="C"><%=getadminRepEditLngStr("DtxtCenter")%></option>
				<option <% If rs("colAlign") = "R" Then %>selected<% End If %> value="R"><%=getadminRepEditLngStr("DtxtRight")%></option>
				</select></td>
				<td width="120">
				<select size="1" style="width: 190px" name="Format<%=rs("ID")%>" id="myFormat" class="input">
				<option <% If rs("colFormat") = "N" Then %>selected<% End If %> value="N"></option>
				<% If rs("colType") = 200 or rs("colType") = 202 Then %>
				<option <% If rs("colFormat") = "EML" Then %>selected<% End If %> value="EML"><%=getadminRepEditLngStr("DtxtEMail")%></option>
				<% End If %>
				<% If rs("colType") = 135 Then %>
				<option <% If rs("colFormat") = "DT2" Then %>selected<% End If %> value="DT2"><%=getadminRepEditLngStr("LtxtShortDate")%></option>
				<option <% If rs("colFormat") = "DT1" Then %>selected<% End If %> value="DT1"><%=getadminRepEditLngStr("LtxtLongDate")%></option>
				<option <% If rs("colFormat") = "DT3" Then %>selected<% End If %> value="DT3"><%=getadminRepEditLngStr("LtxtHourAMPM")%></option>
				<option <% If rs("colFormat") = "DT4" Then %>selected<% End If %> value="DT4"><%=getadminRepEditLngStr("LtxtHour24")%></option>
				<% End If %>
				<% If rs("EnableFormat") Then %>
				<option <% If rs("colFormat") = "R" Then %>selected<% End If %> value="R"><%=getadminRepEditLngStr("DtxtRate")%></option>
				<option <% If rs("colFormat") = "S" Then %>selected<% End If %> value="S"><%=getadminRepEditLngStr("DtxtImport2")%></option>
				<option <% If rs("colFormat") = "P" Then %>selected<% End If %> value="P"><%=getadminRepEditLngStr("DtxtPrice")%></option>
				<option <% If rs("colFormat") = "Q" Then %>selected<% End If %> value="Q"><%=getadminRepEditLngStr("DtxtQty")%></option>
				<option <% If rs("colFormat") = "%" Then %>selected<% End If %> value="%"><%=getadminRepEditLngStr("DtxtPercentage")%></option>
				<option <% If rs("colFormat") = "M" Then %>selected<% End If %> value="M"><%=getadminRepEditLngStr("DtxtMeasure")%></option>
				<option <% If rs("colFormat") = "IB" Then %>selected<% End If %> value="IB"><%=getadminRepEditLngStr("LtxtQtyInUn")%></option>
				<option <% If rs("colFormat") = "IS" Then %>selected<% End If %> value="IS"><%=getadminRepEditLngStr("LtxtQtyOutUn")%></option>
				<% End If %>
				<option <% If rs("colFormat") = "IM" Then %>selected<% End If %> value="IM"><%=getadminRepEditLngStr("LtxtImage")%></option>
				<option <% If rs("colFormat") = "H" Then %>selected<% End If %> value="H"><%=getadminRepEditLngStr("LtxtInvisible")%></option>
				</select></td>
				<td width="80">
				<select size="1" name="Show<%=rs("ID")%>">
				<option value="B" <% If rs("colShow") = "B" Then %>selected<% End If %>><%=getadminRepEditLngStr("LtxtDown")%></option>
				<option value="T" <% If rs("colShow") = "T" Then %>selected<% End If %>><%=getadminRepEditLngStr("LtxtUp")%></option>
				<option value="A" <% If rs("colShow") = "A" Then %>selected<% End If %>><%=getadminRepEditLngStr("LtxtBoth")%></option>
				</select></td>
				<td width="80"><select size="1" style="width: 73px" name="action<%=rs("ID")%>" id="myAction">
				<option <% If rs("colTotal") = "D" Then %>selected<% End If %> value="D"></option>
				<% If rs("EnableFormat") Then %><option <% If rs("colTotal") = "Avg" Then %>selected<% End If %> value="Avg"><%=getadminRepEditLngStr("DtxtAverage")%></option><% End If %>
				<option <% If rs("colTotal") = "Count" Then %>selected<% End If %> value="Count"><%=getadminRepEditLngStr("DtxtCount")%></option>
				<option <% If rs("colTotal") = "Max" Then %>selected<% End If %> value="Max"><%=getadminRepEditLngStr("DtxtMaximum")%></option>
				<option <% If rs("colTotal") = "Min" Then %>selected<% End If %> value="Min"><%=getadminRepEditLngStr("DtxtMinimum")%></option>
				<% If rs("EnableFormat") Then %><option <% If rs("colTotal") = "Sum" Then %>selected<% End If %> value="Sum"><%=getadminRepEditLngStr("DtxtSum")%></option><% End If %>
				</select></td>
				<td class="style1">
				<p class="style1">
				<input <% If Not rs("EnableFormat") Then %>disabled<% End If %> type="checkbox" name="colSum<%=rs("ID")%>" id="colSum<%=rs("ID")%>" class="OptionButton" style="background:background-image"<% If rs("colSum") = "Y" Then %>checked<% End If %> value="Y"><label for="colSum<%=rs("ID")%>"><%=getadminRepEditLngStr("DtxtEnable")%></label></td>
				<td class="style1">
				<input type="checkbox" name="colNB<%=rs("ID")%>" id="colNB<%=rs("ID")%>" class="OptionButton" style="background:background-image"<% If rs("colNB") = "Y" Then %>checked<% End If %> value="Y"><label for="colNB<%=rs("ID")%>"><%=getadminRepEditLngStr("DtxtEnable")%></label></td>
				<td width="58">
				<table border="0" cellspacing="0" id="table11">
					<tr>
						<td>
						<img border="0" id="imgLnk<%=rs("ID")%>" src="images/link_<% If rs("linkType") = "N" Then %>un<% End If %>active.jpg" width="16" height="16"></td>
						<td>
				<input type="button" class="BtnRep" value="<%=getadminRepEditLngStr("DtxtEdit")%>" name="btnEditLink" onclick="javascript:doEditLink(<%=Request("rsIndex")%>,'<%=Replace(myJavascriptEncode(rs("colName")), "'", "\'")%>', 'imgLnk<%=rs("ID")%>', <%=rs("ID")%>);" style="height: 16px;"></td>
					</tr>
				</table>
				</td>
				<td style="width: 18px" class="style1">
				<img src="images/col_link.jpg" border="0" id="colLnk<%=colName%>" <% If rs("ShowLinkVar") = "N" Then %>style="display: none; "<% End If %> alt="<%=getadminRepEditLngStr("LtxtIsLinkCol")%>"></td>
				<td style="width: 18px" class="style1">
				<% If rs("ShowColor") = "Y" Then %><img src="images/col_color.gif" border="0" alt="<%=getadminRepEditLngStr("LtxtIsColorColumn")%>"><% Else %>&nbsp;<% End If %></td>
			</tr>
			<% rs.movenext
			loop
			End If %>
			</table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table9">
			<tr>
				<td width="77">
				<input class="BtnRep" type="submit" value="<%=getadminRepEditLngStr("DtxtSave")%>" name="B2"></td>
				<td><hr size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
	<input type="hidden" name="cmd" value="repTotals">
	<input type="hidden" name="rsIndex" value="<%=rsIndex%>">
	<input type="hidden" name="UserType" value="<%=UserType%>">
</form>
<script language="javascript">
<!--
function updateLinkImg(linkType, linkCols)
{
	try
	{
		if (linkType == 'N') document.getElementById(updLnkImg).src='images/link_unactive.jpg';
		else document.getElementById(updLnkImg).src='images/link_active.jpg';
		
		if (linkCols != '')
		{
			arrCols = linkCols.split('{S}');
			mycols = document.form3.arrColName;
			
			if (mycols.length)
			{
				for (var i = 0;i<mycols.length;i++)
				{
					mycolName = mycols[i].value;
					mycolName = mycolName.substring(1, mycolName.length-1)
					var found = false;
					
					for (var j = 0;j<arrCols.length;j++)
					{
						if (mycolName == arrCols[j])
						{
							found = true;
							break;
						}
					}
					
					if (!found)
					{
						document.getElementById('colLnk' + mycolName).style.display = 'none';
					}
					else
					{
						document.getElementById('colLnk' + mycolName).style.display = '';
					}
				}
			}
			else
			{
				mycols = mycols.substring(1, mycols.length-1)
				var found = false;
				
				for (var j = 0;j<arrCols.length;j++)
				{
					if (mycols == arrCols[j])
					{
						found = true;
						break;
					}
				}
				
				if (!found)
				{
					document.getElementById('colLnk' + mycols).style.display = 'none';
				}
				else
				{
					document.getElementById('colLnk' + mycols).style.display = '';
				}
			}
		}
	}
	catch (ex)
	{
	}
}
var updLnkImg;
function doEditLink(rsIndex, colName, lnkImg, id)
{
	updLnkImg = lnkImg;
	OpenWin = this.open('', "CtrlWindow", "toolbar=no,menubar=no,location=no,scrollbars=yes,resizable=no, width=552,height=256");
	doMyLink('adminRepEditLinks.asp', 'rsIndex='+rsIndex+'&colName=' + colName + '&id=' + id + '&UserType=<%=UserType%>&pop=Y', 'CtrlWindow');
}

function valFrm3()
{
<% If Not clearQry Then %>
foundNumInSale = <%=foundNumInSale%>;
foundNumInBuy = <%=foundNumInBuy%>;
<% End If %>
	<% If cCount = 1 Then %>
	fVal = document.form3.myFormat.value;
	if ((fVal == 'H' || fVal == 'IS' || fVal == 'IB') && document.form3.myAction.selectedIndex != 0)
	{
		errMsg = '<%=getadminRepEditLngStr("LtxtValNoFormat")%> ';
		switch (fVal)
		{
			case "H":
				errMsg += '<%=getadminRepEditLngStr("LtxtInvisible")%>';
				break;
			case "IS":
				errMsg += '<%=Replace(getadminRepEditLngStr("LtxtQtyOutUn"), "'", "\'")%>';
				break;
			case "IB":
				errMsg += '<%=Replace(getadminRepEditLngStr("LtxtQtyInUn"), "'", "\'")%>';
				break;
		}
		alert(errMsg);
		return false;
	}
	else if (fVal == 'IS' && !foundNumInSale)
	{
		alert('<%=getadminRepEditLngStr("LtxtValNoNumInSale")%>');
		return false;
	}
	else if (fVal == 'IB' && !foundNumInBuy)
	{
		alert('<%=getadminRepEditLngStr("LtxtValNoNumInBuy")%>');
		return false;
	}
	<% ElseIf cCount > 1 Then %>
	f = document.form3.myFormat;
	a = document.form3.myAction;
	for (var i = 0;i<f.length;i++)
	{
		fVal = f[i].value;
		if ((fVal == 'H' || fVal == 'IS' || fVal == 'IB') && a[i].selectedIndex != 0)
		{
			errMsg = '<%=getadminRepEditLngStr("LtxtValNoFormat")%> ';
			switch (fVal)
			{
				case "H":
					errMsg += '<%=getadminRepEditLngStr("LtxtInvisible")%>';
					break;
				case "IS":
					errMsg += '<%=Replace(getadminRepEditLngStr("LtxtQtyOutUn"), "'", "\'")%>';
					break;
				case "IB":
					errMsg += '<%=Replace(getadminRepEditLngStr("LtxtQtyInUn"), "'", "\'")%>';
					break;
			}
			alert(errMsg);
			return false;
		}
		else if (fVal == 'IS' && !foundNumInSale)
		{
			alert('<%=getadminRepEditLngStr("LtxtValNoNumInSale")%>');
			return false;
		}
		else if (fVal == 'IB' && !foundNumInBuy)
		{
			alert('<%=getadminRepEditLngStr("LtxtValNoNumInBuy")%>');
			return false;
		}
	}
	<% End If %>
	return true;
}
//-->
</script>
	<% ElseIf Request("repCmd") = "repColor" Then
	rs.close
	sql = "select ColorID, LineID, Alias, colName, " & _
		  "Case colOp " & _
		  "When '=' 	Then N'" & getadminRepEditLngStr("LtxtEqualTo") & "' " & _
		  "When '<>' 	Then N'" & getadminRepEditLngStr("LtxtNotEqual") & "' " & _
		  "When '>' 	Then N'" & getadminRepEditLngStr("LtxtMoreThen") & "' " & _
		  "When '<' 	Then N'" & getadminRepEditLngStr("LtxtLessThen") & "' " & _
		  "When '>=' 	Then N'" & getadminRepEditLngStr("LtxtMoreOrEq") & "' " & _
		  "When '<=' 	Then N'" & getadminRepEditLngStr("LtxtLessOrEq") & "' " & _
		  "When 'N' 	Then N'" & getadminRepEditLngStr("LtxtNull") & "' " & _
		  "When 'NN' 	Then N'" & getadminRepEditLngStr("LtxtNotNull") & "' End colOp, " & _
		  "Case colOpBy When 'F' Then N'" & getadminRepEditLngStr("DtxtField") & "' When 'V' Then N'" & getadminRepEditLngStr("DtxtValue") & "' End colOpBy, colValue, " & _
		  "Case ApplyToRow When 'Y' Then N'" & getadminRepEditLngStr("DtxtYes") & "' When 'N' Then N'" & getadminRepEditLngStr("DtxtNo") & "' End ApplyToRow, " & _
		  "Active, " & _
		  "FontFace, FontSize, ForeColor, FontBold, FontItalic, FontUnderline, FontStrike, FontBlink, FontAlign, BackColor, T0.Ordr, T0.Ordr2 " & _
		  "from OLKRSColors T0 " & _
		  "where rsIndex = " & Request("rsIndex") & " order by T0.Ordr, T0.Ordr2"
	rs.open sql, conn, 3, 1
	
	If Not rs.Eof Then
	%>
	<tr class="TblRepNrm">
		<td>
		<p align="justify"><img src="images/lentes.gif">
		<%=getadminRepEditLngStr("LttlColorNote")%></td>
	</tr>
	<form name="frmAdmColors" action="repSubmit.asp" method="post" onsubmit="return valFrmAdmCol();">
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="tblRepColors">
			<tr class="TblRepTltSub">
				<td width="14"></td>
				<td width="40">#</td>
				<td class="style1"><%=getadminRepEditLngStr("LtxtAlias")%></td>
				<td class="style1"><%=getadminRepEditLngStr("LtxtColumn")%></td>
				<td class="style1"><%=getadminRepEditLngStr("LtxtOperation")%></td>
				<td class="style1"><%=getadminRepEditLngStr("DtxtBy")%></td>
				<td width="116" class="style1"><%=getadminRepEditLngStr("DtxtField")%>/<%=getadminRepEditLngStr("DtxtValue")%></td>
				<td class="style1"><%=getadminRepEditLngStr("DtxtExample")%></td>
				<td class="style1"><%=getadminRepEditLngStr("DtxtLine")%></td>
				<td class="style1"><%=getadminRepEditLngStr("DtxtOrder")%></td>
				<td class="style1"><%=getadminRepEditLngStr("LtxtAlternative")%></td>
				<td class="style1"><%=getadminRepEditLngStr("DtxtActive")%></td>
				<td width="56" colspan="2"></td>
			</tr>
			<% 
			LineNum = 0 'Grupo
			AlterNum = 0 'Linea de Grupo
			LastColorID = -1 'Ultimo Color
			
			do while not rs.eof
			
			If rs("ColorID") <> LastColorID Then
				LineNum = LineNum + 1
				LastColorID = rs("ColorID")
				AlterNum = 0
			Else
				AlterNum = AlterNum + 1
			End If
			
			ColID = "Lst" & rs("ColorID") & "_" & rs("LineID") %>
			<input type="hidden" name="ColID" value="<%=ColID%>">
			<tr class="<% If Request("ColorID") = CStr(rs("ColorID")) and Request("LineID") = CStr(rs("LineID")) Then %>TblRepTbl<% Else %>TblRepNrm<% End If %>">
				<td width="14">
				<a href="javascript:doEditColor(<%=Request("rsIndex")%>,<%=rs("ColorID")%>,<%=rs("LineID")%>,<%=LineNum%>,<%=AlterNum%>);"><img border="0" src="images/<%=Session("rtl")%>flechaselec.gif" width="15" height="13"></a>
				</td>
				<td width="40"><% For i = 0 to AlterNum %>&nbsp;<% Next %><% If AlterNum = 0 Then %><%=LineNum%><% Else %><img src="images/<%=Session("rtl")%>alter.jpg" border="0"><%=AlterNum%><% End if %></td>
				<td width="200">
				<table cellpadding="0" cellspacing="0" border="0" width="100%">
					<tr class="<% If Request("ColorID") = CStr(rs("ColorID")) and Request("LineID") = CStr(rs("LineID")) Then %>TblRepTbl<% Else %>TblRepNrm<% End If %>">
						<td><input class="input" style="width: 100%; " size="20" value="<%=Server.HTMLEncode(rs("Alias"))%>" name="ColAlias<%=ColID%>" id="ColAlias" onkeydown="return chkMax(event, this, 50);" maxlength="50"></td>
						<td width="16"><a href="javascript:doFldTrad('RSColors', 'rsIndex,ColorID,LineID', '<%=Request("rsIndex")%>,<%=rs("ColorID")%>,<%=rs("LineID")%>', 'alterAlias', 'T', null);"><img src="images/trad.gif" alt="<%=getadminRepEditLngStr("DtxtTranslate")%>" border="0"></a></td>
					</tr>
				</table>
				</td>
				<td><%=rs("colName")%>&nbsp;</td>
				<td><%=rs("colOp")%>&nbsp;</td>
				<td><%=rs("colOpBy")%>&nbsp;</td>
				<td width="116"><%=rs("colValue")%>&nbsp;</td>
				<td bgcolor="<%=rs("BackColor")%>"><%=setColFormat(getadminRepEditLngStr("DtxtExample"), rs("FontFace"), rs("FontSize"), rs("ForeColor"), rs("FontBold"), rs("FontItalic"), rs("FontUnderline"), rs("FontStrike"), rs("FontBlink"), rs("FontAlign"))%>&nbsp;</td>
				<td><%=rs("ApplyToRow")%>&nbsp;</td>
				<td align="center">
				<% If AlterNum = 0 Then %><table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td>
						<input id="ColOrdr<%=rs("ColorID")%>" name="ColOrdr<%=rs("ColorID")%>" value="<%=rs("Ordr")%>" size="6" onchange="if(!IsNumeric(this.value) || this.value=='')this.value=<%=rs("Ordr")%>" onfocus="this.select();" onmouseup="event.preventDefault()" maxlength="6"></td>
						<td valign="middle">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td><img src="images/img_nud_up.gif" id="ColOrdr<%=rs("ColorID")%>Up"></td>
							</tr>
							<tr>
								<td><img src="images/spacer.gif"></td>
							</tr>
							<tr>
								<td><img src="images/img_nud_down.gif" id="ColOrdr<%=rs("ColorID")%>Down"></td>
							</tr>
						</table>
						</td>
					</tr>
				</table>
				<script language="javascript">NumUDAttachMin('frmAdmColors', 'ColOrdr<%=rs("ColorID")%>', 'ColOrdr<%=rs("ColorID")%>Up', 'ColOrdr<%=rs("ColorID")%>Down', 1);</script>
				<% End If %></td>
				<td style="text-align: center;">
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td>
						<input id="ColOrdr2<%=ColID%>" name="ColOrdr2<%=ColID%>" value="<%=rs("Ordr2")%>" size="6" onchange="if(!IsNumeric(this.value) || this.value=='')this.value=<%=rs("Ordr2")%>" onfocus="this.select();" onmouseup="event.preventDefault()" maxlength="6"></td>
						<td valign="middle">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td><img src="images/img_nud_up.gif" id="ColOrdr2<%=ColID%>Up"></td>
							</tr>
							<tr>
								<td><img src="images/spacer.gif"></td>
							</tr>
							<tr>
								<td><img src="images/img_nud_down.gif" id="ColOrdr2<%=ColID%>Down"></td>
							</tr>
						</table>
						</td>
					</tr>
				</table>
				<script language="javascript">NumUDAttachMin('frmAdmColors', 'ColOrdr2<%=ColID%>', 'ColOrdr2<%=ColID%>Up', 'ColOrdr2<%=ColID%>Down', 1);</script>
				</td>
				<td style="text-align: center;"><input type="checkbox" class="noborder"  name="ColActive<%=ColID%>" value="Y" <% If rs("Active") = "Y" Then %>checked<% End If %>></td>
				<td width="20">
				<% If AlterNum = 0 Then %>
				<input type="button" name="btnAddAlter" value="+ <%=getadminRepEditLngStr("LtxtAlternative")%>" onclick="javascript:addAlterColor(<%=Request("rsIndex")%>,<%=rs("ColorID")%>,<%=LineNum%>);">
				<% End If %></td>
				<td width="16">
				<img border="0" src="images/remove.gif" style="cursor: hand" onclick="javascript:doDelColor(<%=Request("rsIndex")%>,<%=rs("ColorID")%>,<%=rs("LineID")%>);"></td>
			</tr>
			<% rs.movenext
			loop %>
			</table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table9">
			<tr>
				<td width="77">
				<input class="BtnRep" type="submit" value="<%=getadminRepEditLngStr("DtxtSave")%>" name="btnSave"></td>
				<td><hr size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
	<input type="hidden" name="cmd" value="AdmColors">
	<input type="hidden" name="rsIndex" value="<%=rsIndex%>">
	<input type="hidden" name="UserType" value="<%=UserType%>">
	</form>
	<% End If
	sql = getGenQry
	set rs = conn.execute(sql)
	sql = ""
	Dim ArrCol()
	Redim ArrCol(rs.Fields.Count-1)
	Dim ArrColItm(1)
	For i = 0 to rs.Fields.count -1
		myColTypeVal = getColTypeVal(rs(i).Type)
		If myColTypeVal <> "U" Then
			ArrColItm(0) = rs.Fields(i).Name
			ArrColItm(1) = rs.Fields(i).Name & "{|}" & myColTypeVal
			ArrCol(i) = ArrColItm
		End If
	next
	rs.close
	If Request("ColorID") <> "" and Request("LineID") <> "" Then
		sql = "select T0.Alias, T0.colName, T0.colType, T0.colOp, T0.colOpBy, T0.colValue, T0.colValDate, T0.FontFace, IsNull(T0.FontSize, 1) FontSize,  " & _
		"ForeColor, BackColor, FontBold, FontItalic, FontUnderline, FontStrike, FontBlink, FontAlign, ApplyToRow, ApplyToCol, T0.Active, T0.Ordr, T0.Ordr2 " & _
		"from OLKRSColors T0 " & _
		"where rsIndex = " & Request("rsIndex") & " and ColorID = " & Request("ColorID") & " and LineID = " & Request("LineID")
		set rs = conn.execute(sql)
		Alias = rs("Alias")
		colName = rs("colName")
		colType = rs("colType")
		colOp = rs("colOp")
		colOpBy = rs("colOpBy")
		colValue = rs("colValue")
		colValDate = rs("colValDate")
		If rs("FontFace") <> "" Then FontFace = rs("FontFace") Else FontFace = "&nbsp;"
		If rs("FontSize") <> "" Then FontSize = rs("FontSize") Else FontSize = "&nbsp;"
		ForeColor = rs("ForeColor")
		BackColor = rs("BackColor")
		FontBold = rs("FontBold")
		FontItalic = rs("FontItalic")
		FontUnderline = rs("FontUnderline")
		FontStrike = rs("FontStrike")
		FontBlink = rs("FontBlink")
		FontAlign = rs("FontAlign")
		ApplyToRow = rs("ApplyToRow")
		ApplyToCol = rs("ApplyToCol")
		Active = rs("Active")
		Ordr = rs("Ordr")
		Ordr2 = rs("Ordr2")
		
		LineNum = Request("LineNum")
		AlterLineNum = Request("AlterNum")
	ElseIf Request("ColorID") <> "" Then
		sql = "select top 1 ApplyToRow, ApplyToCol, Ordr, " & _
		"(select Max(Ordr2)+1 from OLKRSColors where rsIndex = T0.rsIndex and ColorID = T0.ColorID) Ordr2 " & _
		"from OLKRSColors T0 " & _
		"where rsIndex = " & Request("rsIndex") & " and ColorID = " & Request("ColorID") & " " & _
		"order by Ordr2 asc"
		set rs = conn.execute(sql)
		ApplyToRow = rs("ApplyToRow")
		ApplyToCol = rs("ApplyToCol")
		colOpBy = "F"
		FontFace = "&nbsp;"
		FontSize = "1"
		Ordr = rs("Ordr")
		Ordr2 = rs("Ordr2")
		
		LineNum = Request("LineNum")
		AlterLineNum = Request("AlterNum")
	Else
		sql = "select IsNull((select Max(Ordr)+1 from OLKRSColors where rsIndex = " & Request("rsIndex") & "), 0) Ordr"
		set rs = conn.execute(sql)
		colOpBy = "F"
		FontFace = "&nbsp;"
		FontSize = "1"
		Alias = ""
		colValue = ""
		Ordr = rs("Ordr")
		Ordr2 = 0
		
		LineNum = LineNum+1
		AlterNum = 0 %>
	<input type="hidden" name="AliasTrad">
	<% End If %>
	<form name="frmAddEditCol" action="repSubmit.asp" method="POST" onsubmit="return valFrmAddEditCol();">
	<tr class="TblRepTlt" id="editColor">
		<td>
 		<% If Request("ColorID") = "" or Request("AlterOf") <> "" Then %><%=getadminRepEditLngStr("DtxtAdd")%>&nbsp;<% ElseIf Request("LineID") <> "" Then %><%=getadminRepEditLngStr("DtxtEdit")%>&nbsp;<% End If %><%=getadminRepEditLngStr("DtxtColor")%><% If Request("ColorID") <> "" or Request("AlterOf") <> "" Then %>&nbsp;<% If Request("AlterOf") <> "" Then %><%=getadminRepEditLngStr("LtxtAlterOf")%>&nbsp;<% End If %>#<%=Request("LineNum")%><% If Request("AlterNum") > 0 Then Response.Write "-" & Request("AlterNum") %><% End If %></td>
	</tr>
	<tr class="TblRepNrm">
		<td>
		<p align="justify"><img src="images/lentes.gif"><%=getadminRepEditLngStr("LtxtAddColorNote")%></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" id="table8" style="width: 100%; ">
			<tr>
				<td colspan="2" class="style2">
					
					<%=getadminRepEditLngStr("LtxtConf")%></td>
				<td colspan="2" class="style2">
					<%=getadminRepEditLngStr("LtxtFormat")%></td>
			</tr>
			<tr>
				<td class="TblRepTlt">
										<%=getadminRepEditLngStr("LtxtAlias")%></td>
					<td class="TblRepTbl">
					<table cellpadding="0" cellspacing="0" border="0" style="width: 100%;">
						<tr>
							<td><input type="text" name="Alias" size="50" style="width: 100%; " maxlength="50" value="<%=Server.HTMLEncode(Alias)%>">
							</td>
							<td width="16"><a href="javascript:doFldTrad('RSColors', 'rsIndex,ColorID,LineID', '<%=Request("rsIndex")%>,<%=Request("ColorID")%>,<%=Request("LineID")%>', 'alterAlias', 'T', <% If Request("LineID") <> "" Then %>null<% Else %>document.frmAddEditCol.AliasTrad<% End If %>);"><img src="images/trad.gif" alt="<%=getadminRepEditLngStr("DtxtTranslate")%>" border="0"></a></td>
						</tr>
					</table>
					</td>

				<td class="TblRepTlt">
					<% If Request("AlterNum") = 0 and Request("AlterOf") <> "Y" Then %><%=getadminRepEditLngStr("DtxtOrder")%><% Else %><%=getadminRepEditLngStr("LtxtAlternative")%><% End If %></td>
				<td class="TblRepTbl">
				<% 
				If Request("AlterNum") = 0 and Request("AlterOf") <> "Y" Then
					HideID = "Ordr2"
					HideVal = Ordr2
					
					ShowID = "Ordr"
					ShowVal = Ordr
				Else
					HideID = "Ordr"
					HideVal = Ordr
					
					ShowID = "Ordr2"
					ShowVal = Ordr2
				End If %>
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td>
						<input id="<%=ShowID%>" name="<%=ShowID%>" value="<%=ShowVal%>" size="6" onchange="if(!IsNumeric(this.value) || this.value=='')this.value=<%=rs("Ordr")%>" onfocus="this.select();" onmouseup="event.preventDefault()" maxlength="6"></td>
						<td valign="middle">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td><img src="images/img_nud_up.gif" id="<%=ShowID%>Up"></td>
							</tr>
							<tr>
								<td><img src="images/spacer.gif"></td>
							</tr>
							<tr>
								<td><img src="images/img_nud_down.gif" id="<%=ShowID%>Down"></td>
							</tr>
						</table>
						</td>
					</tr>
				</table>
				<script language="javascript">NumUDAttachMin('frmAddEditCol', '<%=ShowID%>', '<%=ShowID%>Up', '<%=ShowID%>Down', 1);</script>
				<input type="hidden" name="<%=HideID%>" value="<%=HideVal%>"></td>
			</tr>
			<tr>
				<td class="TblRepTlt">
					<span id="txtOpByFld0" <% If colOpBy = "V" Then %> style="display: none"<% End If %>>
					<%=getadminRepEditLngStr("DtxtField")%></span></td>
				<td class="TblRepTbl">
					<select size="1" name="colName" onchange="javascript:ChangeColName(this.value);">
					<% For i = 0 to UBound(ArrCol) %>
					<option <% If colName = ArrCol(i)(0) Then %>selected<% End If %> value="<%=myHTMLEncode(ArrCol(i)(1))%>"><%=myHTMLEncode(ArrCol(i)(0))%></option>
					<% Next %>
					</select></td>
				<td class="TblRepTlt">
										<%=getadminRepEditLngStr("LtxtFont")%></td>
				<td class="TblRepTbl">
					<table border="0" cellspacing="0" width="120" id="table1" cellpadding="0" class="TblCombo">
						<tr>
							<td style="cursor: default; font-size: 10px; font-weight: bold" onclick="return !showSelectFont(imgCmbFontDown, document.frmAddEditCol.FontFace, event);"><span id="txtSelFontSample"><font face="<%=FontFace%>"><%=FontFace%></font></span></td>
							<td width="12"><img src="images_picker/select_arrow_small.gif" id="imgCmbFontDown" onmouseover="this.src='images_picker/select_arrow_over_small.gif'" onmouseout="this.src='images_picker/select_arrow_small.gif'" onclick="return !showSelectFont(this, document.frmAddEditCol.FontFace, event);"></td>
						</tr>
						<input type="hidden" name="FontFace" value="<%=FontFace%>">
					</table></td>
			</tr>
			<tr>
				<td class="TblRepTlt">
					<%=getadminRepEditLngStr("LtxtOperation")%></td>
				<td class="TblRepTbl">
					<select size="1" name="colOp" onchange="javascript:ChangeOp(this.value);">
					<option value="=" <% If colOp = "=" Then %>selected<% End If %>>
					<%=getadminRepEditLngStr("LtxtEqualTo")%></option>
					<option value="<>" <% If colOp = "<>" Then %>selected<% End If %>>
					<%=getadminRepEditLngStr("LtxtNotEqual")%></option>
					<option value=">" <% If colOp = ">" Then %>selected<% End If %>>
					<%=getadminRepEditLngStr("LtxtMoreThen")%></option>
					<option value="<" <% If colOp = "<" Then %>selected<% End If %>>
					<%=getadminRepEditLngStr("LtxtLessThen")%></option>
					<option value=">=" <% If colOp = ">=" Then %>selected<% End If %>>
					<%=getadminRepEditLngStr("LtxtMoreOrEq")%></option>
					<option value="<=" <% If colOp = "<=" Then %>selected<% End If %>>
					<%=getadminRepEditLngStr("LtxtLessOrEq")%></option>
					<option value="N" <% If colOp = "N" Then %>selected<% End If %>>
					<%=getadminRepEditLngStr("LtxtNull")%></option>
					<option value="NN" <% If colOp = "NN" Then %>selected<% End If %>>
					<%=getadminRepEditLngStr("LtxtNotNull")%></option>
					</select></td>
				<td class="TblRepTlt">
										<%=getadminRepEditLngStr("DtxtSize")%></td>
				<td class="TblRepTbl">
					<table border="0" cellspacing="0" width="120" id="table1" cellpadding="0" class="TblCombo">
						<tr>
							<td style="cursor: default; font-size: 10px; font-weight: bold" onclick="return !showSelectFontSize(imgCmbFontSizeDown, document.frmAddEditCol.FontSize, event);"><font face="Verdana" size="1"><span id="txtSelFontSizeSample"><%=FontSize%></span></font></td>
							<td width="12"><img src="images_picker/select_arrow_small.gif" id="imgCmbFontSizeDown" onmouseover="this.src='images_picker/select_arrow_over_small.gif'" onmouseout="this.src='images_picker/select_arrow_small.gif'" onclick="return !showSelectFontSize(this, document.frmAddEditCol.FontSize, event);"></td>
						</tr>
						<input type="hidden" name="FontSize" value="<%=FontSize%>">
					</table>
					</td>
			</tr>
			<tr>
				<td class="TblRepTlt">
					<%=getadminRepEditLngStr("DtxtBy")%></td>
				<td class="TblRepTbl">
					<select size="1" <% If colOp = "N" or colOp = "NN" Then %>disabled<% End If %> name="colOpBy" onchange="ChangeOpBy(this.value);">
					<option value="F"><%=getadminRepEditLngStr("DtxtField")%></option>
					<option value="V" <% If colOpBy = "V" Then %>selected<% End If %>>
					<%=getadminRepEditLngStr("DtxtValue")%></option>
					</select></td>
				<td class="TblRepTlt">
										<%=getadminRepEditLngStr("LtxtForeColor")%></td>
				<td class="TblRepTbl">
					<table cellpadding="0" cellpadding="2" border="0">
						<tr>
							<td style="width: 74px">
							<table cellpadding="0" cellspacing="0" border="0" bgcolor="#D9F0FD" style="border: 1px solid">
								<tr>
									<td style="border-left-style: none; border-left-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium">
										<input type="text" readonly name="ForeColor" size="9" style="cursor: default" maxlength="7" value="<%=ForeColor%>" onclick="showColorPicker(btnChangeForeColor,this,ForeColorSample)"></td>
									<td style="border-right-style: none; border-right-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium"><img src="images_picker/select_arrow_small.gif" onmouseover="this.src='images_picker/select_arrow_over_small.gif'" onmouseout="this.src='images_picker/select_arrow_small.gif'" id="btnChangeForeColor" onclick="showColorPicker(this,document.frmAddEditCol.ForeColor,ForeColorSample)"></td>
								</tr>
							</table>
							</td>
							<td width="46" bgcolor="<%=ForeColor%>" style="border: 1px solid; font-size: 10px" id="ForeColorSample">
							&nbsp;</td>
							<td><a href="javascript:clearColor(document.frmAddEditCol.ForeColor, ForeColorSample);"><img border="0" src="images/remove.gif"></a></td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td class="TblRepTlt">
					<span id="txtOpByVal" <% If colOpBy = "F" Then %>style="display: none"<% End If %>>
					<%=getadminRepEditLngStr("DtxtValue")%></span><span id="txtOpByFld" <% If colOpBy = "V" Then %> style="display: none"<% End If %>><%=getadminRepEditLngStr("DtxtField")%></span>&nbsp;</td>
				<td class="TblRepTbl">
					<table border="0" id="tblValDat" cellspacing="0" cellpadding="0" style="<% If colOpBy = "F" or colOpBy = "V" and colType <> "D" Then %>display: none<% End If %>">
						<tr>
							<td><img border="0" src="images/cal.gif" id="btnValDatImg" width="16" height="16" style="float:left;padding-left:1px;padding-top:1px"></td>
							<td>
							<input type="text" <% If colOp = "N" or colOp = "NN" Then %>disabled<% End If %> readonly name="colValDat" id="colValDat" size="10" value="<%=FormatDate(colValDate, False)%>" onclick="btnValDatImg.click()"></td>
							<td><img border="0" src="images/remove.gif" style="cursor: hand" onclick="javascript:document.frmAddEditCol.colValDat.value='';"></td>
						</tr>
					</table>
					<script language="javascript">
								Calendar.setup({
								    inputField     :    "colValDat",     // id of the input field
								    ifFormat       :    "<%=GetCalendarFormatString%>",      // format of the input field
								    button         :    "btnValDatImg",  // trigger for the calendar (button ID)
								    align          :    "Bl",           // alignment (defaults to "Bl")
								    singleClick    :    true
								});
					</script>
					<input type="text" <% If colOp = "N" or colOp = "NN" Then %>disabled<% End If %> name="colValVal" size="20" style="<% If colOpBy = "F" or colType = "D" Then %>display: none<% End If %>" value="<% If colOpBy = "V" and Not IsNull(colValue) Then Response.Write Server.HTMLEncode(colValue)%>" onchange="valValVal(this);">
					<select <% If colOp = "N" or colOp = "NN" Then %>disabled<% End If %> size="1" name="colValCol" style="<% If colOpBy = "V" Then %>display: none<% End If %>">
					<option value=""></option>
					<% For i = 0 to UBound(ArrCol) %>
					<option <% If colOpBy = "F" and colValue = ArrCol(i)(0) Then %>selected<% End If %> value="<%=myHTMLEncode(ArrCol(i)(1))%>"><%=myHTMLEncode(ArrCol(i)(0))%></option>
					<% Next %>
					</select></td>
				<td class="TblRepTlt">
										<%=getadminRepEditLngStr("LtxtBackColor")%></td>
				<td class="TblRepTbl">
					<table cellpadding="0" cellpadding="2" border="0">
						<tr>
							<td>
							<table cellpadding="0" cellspacing="0" border="0" bgcolor="#D9F0FD" style="border: 1px solid">
								<tr>
									<td style="border-left-style: none; border-left-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium">
										<input type="text" readonly name="BackColor" size="9" style="cursor: default" maxlength="7" value="<%=BackColor%>" onclick="showColorPicker(btnChangeBackColor,this,BackColorSample)"></td>
									<td style="border-right-style: none; border-right-width: medium; border-top-style: none; border-top-width: medium; border-bottom-style: none; border-bottom-width: medium"><img src="images_picker/select_arrow_small.gif" onmouseover="this.src='images_picker/select_arrow_over_small.gif'" onmouseout="this.src='images_picker/select_arrow_small.gif'" id="btnChangeBackColor" onclick="showColorPicker(this,document.frmAddEditCol.BackColor,BackColorSample)"></td>
								</tr>
							</table>
							</td>
							<td width="46" bgcolor="<%=BackColor%>" style="border: 1px solid; font-size: 10px" id="BackColorSample">
							&nbsp;</td>
							<td><a href="javascript:clearColor(document.frmAddEditCol.BackColor, BackColorSample);"><img border="0" src="images/remove.gif"></a></td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td rowspan="3" valign="top" class="TblRepTlt">
					<%=getadminRepEditLngStr("LtxtApplyTo")%></td>
				<td class="TblRepTbl">
					<input type="radio" <% If Request("AlterOf") <> "" or Request("LineID") <> "" and Request("AlterNum") <> 0 Then %>disabled<% End If %> value="F" <% If ApplyToRow = "N" and IsNull(ApplyToCol) or ApplyToCol = "" Then %>checked<% End If %> name="ApplyTo" id="AppyToF" onclick="javascript:ApplyToCol.disabled=true;" class="OptionButton" checked><label for="AppyToF"><%=getadminRepEditLngStr("DtxtField")%></label></td>
					<td class="TblRepTlt">
				<%=getadminRepEditLngStr("DtxtAlignment")%>&nbsp;</td>
					<td class="TblRepTbl">
				<select size="1" name="FontAlign">
				<option></option>
				<option <% If FontAlign = "L" Then %>selected<% End If %> value="L"><%=getadminRepEditLngStr("DtxtLeft")%></option>
				<option <% If FontAlign = "C" Then %>selected<% End If %> value="C"><%=getadminRepEditLngStr("DtxtCenter")%></option>
				<option <% If FontAlign = "R" Then %>selected<% End If %> value="R"><%=getadminRepEditLngStr("DtxtRight")%></option>
				</select></td>
			</tr>
			<tr>
				<td class="TblRepTbl">
					<input type="radio" <% If Request("AlterOf") <> "" or Request("LineID") <> "" and Request("AlterNum") <> 0 Then %>disabled<% End If %> name="ApplyTo" <% If ApplyToRow = "Y" Then %>checked<% End If %> value="R" id="ApplyToR" onclick="javascript:ApplyToCol.disabled=true;" class="OptionButton"><label for="ApplyToR"><%=getadminRepEditLngStr("LtxtRow")%></label></td>
				<td rowspan="5" valign="top" class="TblRepTlt">
					<%=getadminRepEditLngStr("LtxtFormat")%></td>
				<td class="TblRepTbl">
					<input style="background:background-image" class="OptionButton" type="checkbox" name="FontBold" <% If FontBold = "Y" Then %>checked<% End If %> value="Y" id="FontBold"><b><label for="FontBold"><%=getadminRepEditLngStr("LtxtBold")%></label></b></td>
			</tr>
			<tr>
				<td class="TblRepTbl">
					<input type="radio" <% If Request("AlterOf") <> "" or Request("LineID") <> "" and Request("AlterNum") <> 0 Then %>disabled<% End If %> name="ApplyTo" <% If Not IsNull(ApplyToCol) and ApplyToCol <> "" Then %>checked<% End If %> value="A" id="ApplyToA" onclick="javascript:ApplyToCol.disabled=false;" class="OptionButton"><label for="ApplyToA"><%=getadminRepEditLngStr("LtxtOther")%></label>
					<select size="1" name="ApplyToCol" id="ApplyToCol" <% If (IsNull(ApplyToCol) or ApplyToCol = "") or (Request("AlterOf") <> "" or Request("LineID") <> "" and Request("AlterNum") <> 0) Then %>disabled<% End If %>>
					<% For i = 0 to UBound(ArrCol) %>
					<option <% If ApplyToCol = ArrCol(i)(0) Then %>selected<% End If %> value="<%=myHTMLEncode(ArrCol(i)(0))%>"><%=myHTMLEncode(ArrCol(i)(0))%></option>
					<% Next %>
					</select></td>
				<td class="TblRepTbl">
					<input style="background:background-image" class="OptionButton" type="checkbox" name="FontItalic" <% If FontItalic = "Y" Then %>checked<% End If %> value="Y" id="FontItalic"><i><label for="FontItalic"><%=getadminRepEditLngStr("LtxtItalic")%></label></i></td>
			</tr>
			<tr>
				<td rowspan="3" valign="top" class="TblRepTlt">
					&nbsp;</td>
				<td class="TblRepTbl" rowspan="3" valign="top">
					<input style="background:background-image" class="OptionButton" type="checkbox" name="Active" <% If Active = "Y" Then %>checked<% End If %> value="Y" id="Active"><label for="Active"><%=getadminRepEditLngStr("DtxtActive")%></label></td>
				<td class="TblRepTbl">
					<input style="background:background-image" class="OptionButton" type="checkbox" name="FontUnderline" <% If FontUnderline = "Y" Then %>checked<% End If %> value="Y" id="FontUnderline"><u><label for="FontUnderline"><%=getadminRepEditLngStr("LtxtUnderline")%></label></u></td>
			</tr>
			<tr>
				<td class="TblRepTbl">
					<input style="background:background-image" class="OptionButton" type="checkbox" name="FontStrike" <% If FontStrike = "Y" Then %>checked<% End If %> value="Y" id="FontStrike"><strike><label for="FontStrike"><%=getadminRepEditLngStr("LtxtScratched")%></label></strike></td>
			</tr>
			<tr>
				<td class="TblRepTbl">
					<input style="background:background-image" class="OptionButton" type="checkbox" name="FontBlink" <% If FontBlink = "Y" Then %>checked<% End If %> value="Y" id="FontBlink"><blink><label for="FontBlink"><%=getadminRepEditLngStr("LtxtBlink")%></label></blink></td>
			</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table9">
			<tr>
				<td width="77">
				<input class="BtnRep" type="submit" value="<%=getadminRepEditLngStr("DtxtApply")%>" name="btnApply"></td>
				<td width="77">
				<input class="BtnRep" type="submit" value="<%=getadminRepEditLngStr("DtxtSave")%>" name="btnSave"></td>
				<td><hr size="1"></td>
				<td width="77">
				<input type="button" class="BtnRep" value="<%=getadminRepEditLngStr("DtxtCancel")%>" name="btnCancel" onclick="javascript:if(confirm('<%=getadminRepEditLngStr("DtxtConfCancel")%>'))window.location.href='adminRepEdit.asp?rsIndex=<%=Request("rsIndex")%>&repCmd=repColor&#editColor'"></td>
			</tr>
		</table>
		</td>
	</tr>
	<input type="hidden" name="cmd" value="repColor">
	<input type="hidden" name="rsIndex" value="<%=rsIndex%>">
	<input type="hidden" name="UserType" value="<%=UserType%>">
	<input type="hidden" name="ColorID" value="<%=Request("ColorID")%>">
	<input type="hidden" name="LineID" value="<%=Request("LineID")%>">
	<input type="hidden" name="AlterOf" value="<%=Request("AlterOf")%>">
	<input type="hidden" name="LineNum" value="<%=LineNum%>">
	<input type="hidden" name="AlterNum" value="<%=AlterNum%>">
	</form>
	<% End If %>
</table>
<% If Request("repCmd") = "repColor" Then %>
<script language="javascript" src="js_font_picker.js"></script>
<script language="javascript">
var txtValNumVal 	= '<%=getadminRepEditLngStr("DtxtValNumVal")%>';
var txtValAlias 	= '<%=getadminRepEditLngStr("LtxtValAlias")%>';
var txtValOpValue 	= '<%=getadminRepEditLngStr("LtxtValOpValue")%>';
var txtValFldOp		= '<%=getadminRepEditLngStr("LtxtValFldOp")%>';
var txtValFldTypes	= '<%=getadminRepEditLngStr("LtxtValFldTypes")%>';
var txtConfDelColor = '<%=getadminRepEditLngStr("LtxtConfDelColor")%>';
var txtValDelColor1 = '<%=getadminRepEditLngStr("LtxtValDelColor1")%>';
var txtValDelColor2 = '<%=getadminRepEditLngStr("LtxtValDelColor2")%>';
</script>
<script language="javascript" src="adminRepEdit.js"></script>
<% End If %>
<script language="javascript">
function valFrmVar()
{
	var varDefBy = getDefBy();
	var varType = document.form2.varType.value;
	if (document.form2.varName.value == '')
	{
		alert('<%=getadminRepEditLngStr("LtxtValVarNam")%>');
		document.form2.varName.focus();
		return false;
	}
	else if (document.form2.varVar.value == '')
	{
		alert('<%=getadminRepEditLngStr("LtxtValVariable")%>');
		document.form2.varVar.focus();
		return false;
	}
	else if (document.form2.varDataType.value == 'nvarchar' && document.form2.varMaxChar.value == '')
	{
		alert('<%=getadminRepEditLngStr("LtxtValMaxChar")%>');
		document.form2.varMaxChar.focus();
		return false;
	}
	else if (document.form2.varQuery.value == '' && (document.form2.varType.value == 'Q' || document.form2.varType.value == 'L' ||
														document.form2.varType.value == 'DD' || document.form2.varType.value == 'CL'))
	{
		alert('<%=getadminRepEditLngStr("LtxtValQryValues")%>');
		document.form2.varQuery.focus();
		return false;
	}
	else if (document.form2.valVarQuery.value == 'Y')
	{
		alert('<%=getadminRepEditLngStr("LtxtValQryVal")%>');
		document.form2.btnVerfyVar.focus();
		return false;
	}
	else if (document.form2.valVarVar.value == 'Y')
	{
		alert('<%=getadminRepEditLngStr("LtxtValVarVar")%>');
		document.form2.btnVerfyVarVar.focus();
		return false;
	}
	else if (varDefBy == 'V' && varType == 'DP' && document.form2.varDefValDate.value == '')
	{
		alert('<%=getadminRepEditLngStr("LtxtValDefVarDat")%>');
		document.form2.varDefValDate.focus();
		return false;
	}
	else if (varDefBy == 'V' && varType != 'DP' && document.form2.varDefValValue.value == '')
	{
		alert('<%=getadminRepEditLngStr("LtxtValDefVarVal")%>');
		document.form2.varDefValValue.focus();
		return false;
	}
	else if (varDefBy == 'Q' && document.form2.varDefValQuery.value == '')
	{
		alert('<%=getadminRepEditLngStr("LtxtValDefVarQry")%>');
		document.form2.varDefValQuery.focus();
		return false;
	}
	else if (varDefBy == 'Q' && document.form2.valDefValQuery.value == 'Y')
	{
		alert('<%=getadminRepEditLngStr("LtxtValDefVarVerfy")%>');
		document.form2.btnDefValQuery.focus();
		return false;
	}
	return true;
}

function changeType(dType)
{
	if (dType == 'Q' || dType == 'L' || dType == 'DD' || dType == 'CL')
	{
		document.form2.varQuery.style.backgroundColor = "";
		document.form2.varQuery.disabled = false;
		document.form2.varQueryField.disabled = !(document.form2.varQueryBy[0].checked && dType == 'Q');
		enableBranchIndex(!document.form2.varQueryBy[0].checked);
	}
	else
	{
		document.form2.varQuery.style.backgroundColor = "#CCCCCC";
		document.form2.varQuery.disabled = true;
		document.form2.varQueryField.disabled = true;
		enableBranchIndex(false);
	}
	
	if (dType != 'DP' && document.form2.varDataType.selectedIndex == 1)
	{
		document.form2.varDataType.selectedIndex = 0;
		enableVarVerfy();
	}
	else if (dType == 'DP' && document.form2.varDataType.selectedIndex != 1)
	{
		document.form2.varDataType.selectedIndex = 1;
		enableVarVerfy();
	}
	
	changeDefVarBy(getDefBy());
}
function enableBranchIndex(enable)
{
	<% If enabledBaseVars Then %>
		if (document.form2.baseIndex.length)
		{
			for (var i = 0;i<document.form2.baseIndex.length;i++)
			{
				document.getElementById('baseVars' + document.form2.baseIndex[i].value).disabled = !enable;
			}
		}
		else
		{
			document.getElementById('baseVars' + document.form2.baseIndex.value).disabled = !enable;
		}
	<% End If %>
}
function getDefBy()
{
	var retVal = '';
	for (var i = 0;i<document.form2.varDefBy.length;i++)
	{
		if (document.form2.varDefBy[i].checked)
		{
			retVal = document.form2.varDefBy[i].value;
			break;
		}
	}
	return retVal;
}
function chkTypeDate(val)
{
	if (val.value == 'datetime' && document.form2.varType.value != 'DP')
	{
		alert('<%=getadminRepEditLngStr("LtxtValDatFormat")%>');
		val.selectedIndex = 0;
		return;
	}
	else if (val.value != 'datetime' && document.form2.varType.value == 'DP')
	{
		alert('<%=getadminRepEditLngStr("LtxtValTypeDatFormat")%>')
		val.selectedIndex = 1;
		return;
	}
	enableVarVerfy();
}
function enableVarVerfy()
{
	if (document.form2.varQuery.disabled || !document.form2.varQuery.disabled && document.form2.varQuery.value == '')
	{
		document.form2.btnVerfyVar.src='images/btnValidateDis.gif';
		document.form2.btnVerfyVar.style.cursor = '';
		document.form2.valVarQuery.value = 'N';
	}
	else
	{
		document.form2.btnVerfyVar.src='images/btnValidate.gif';
		document.form2.btnVerfyVar.style.cursor = 'hand';
		document.form2.valVarQuery.value = 'Y';
	}
	
	if (document.form2.varDefValQuery.disabled || !document.form2.varDefValQuery.disabled && document.form2.varDefValQuery.value == '')
	{
		document.form2.btnDefValQuery.src='images/btnValidateDis.gif';
		document.form2.btnDefValQuery.style.cursor = '';
		document.form2.valDefValQuery.value = 'N';
	}
	else
	{
		document.form2.btnDefValQuery.src='images/btnValidate.gif';
		document.form2.btnDefValQuery.style.cursor = 'hand';
		document.form2.valDefValQuery.value = 'Y';
	}
}
</script>
<% If Request("repCmd") = "repColor" Then %>
<script language="javascript">setInterval('blinkIt()',1000);</script><% End If %><!--#include file="repBottom.inc" --><!--#include file="getColType.inc"--><% 

Function setColFormat(ByVal colValue, ByVal FontFace, ByVal FontSize, ByVal ForeColor, ByVal FontBold, ByVal FontItalic, ByVal FontUnderline, ByVal FontStrike, ByVal FontBlink, ByVal FontAlign)
	retVal = "<font"
	If Not IsNull(FontFace) Then retVal = retVal & " face=""" & FontFace & """"
	If Not IsNull(FontSize) Then retVal = retVal & " size=""" & FontSize & """"
	If Not IsNull(ForeColor) Then retVal = retVal & " color=""" & ForeColor & """"
	retVal = retVal & ">"
	If retVal = "<font>" Then retVal = ""
	noCloseFont = True
	If FontBold = "Y" Then retVal = retVal & "<b>"
	If FontItalic = "Y" Then retVal = retVal & "<i>"
	If FontUnderline = "Y" Then retVal = retVal & "<u>"
	If FontStrike = "Y" Then retVal = retVal & "<s>"
	If FontBlink = "Y" Then retVal = retVal & "<blink>"
	Select Case FontAlign
		Case "L"
			retVal = retVal & "<p align=""left"">"
		Case "C"
			retVal = retVal & "<p align=""center"">"
		Case "R"
			retVal = retVal & "<p align=""right"">"
	End Select
	retVal = retVal & colValue
	If FontBlink = "Y" Then retVal = retVal & "</blink>"
	If FontStrike = "Y" Then retVal = retVal & "</s>"
	If FontUnderline = "Y" Then retVal = retVal & "</u>"
	If FontItalic = "Y" Then retVal = retVal & "</i>"
	If FontBold = "Y" Then retVal = retVal & "</b>"
	If Not noCloseFont  Then retVal = retVal & "</font>"
	setColFormat = retVal
End Function

Function getRSVariables(ByVal baseIndex)
	strRSVariables = ""
	Select Case Request("UserType") 
		Case "C"
			strRSVariables = "declare @CardCode nvarchar(15) set @CardCode = '' "
		Case "V"
			strRSVariables = "declare @SlpCode int set @SlpCode = -1 "
	End Select
	If Request("baseIndex") <> "" Then %>
	<!--#include file="repVars.inc"-->
<%	
		sql2 = "select '@' + varVar varVar, varDataType, varMaxChar, DefValBy, DefValDate, DefValValue, OLKCommon.dbo.DBOLKGetRSVarBaseIndex" & Session("ID") & "(rsIndex, varIndex) BaseIndex from OLKRSVars where rsIndex = " & Request("rsIndex") & " and varIndex in (" & baseIndex & ")"
		set rsVar = conn.execute(sql2)
		do while not rsVar.eof
			If rsVar("varDataType") = "nvarchar" Then 
				MaxVar = "(" & rsVar("varMaxChar") & ")"
			ElseIf rsVar("varDataType") = "numeric" Then
				MaxVar = "(19,6)"
			Else
				MaxVar = ""
			End If
			strRSVariables = strRSVariables & "declare " & rsVar("varVar") & " " & rsVar("varDataType") & " " & MaxChar & " "
			
			strRSVariables = strRSVariables & "set " & myVar & " = "
			Select Case rsVar("DefValBy")
				Case "V"
					Select Case rsVar("varDataType")
						Case "int", "numeric"
							strRSVariables = strRSVariables & rsVar("DefValValue") & " "
						Case "datetime"
							strRSVariables = strRSVariables & "Convert(datetime,'" & SaveSqlDate(FormatDate(rsVar("DefValDate"), False)) & "',120) "
						Case "nvarchar"
							strRSVariables = strRSVariables & "N'" & rsVar("DefValValue") & "' "
					End Select
				Case "Q"
					set rsVal = Server.CreateObject("ADODB.RecordSet")
					sqlVal = getRSVariables(rsVar("BaseIndex")) & " " & rsVar("DefValValue")
					set rsVal = conn.execute(sqlVal)
					Select Case rsVar("varDataType")
						Case "int", "numeric"
							strRSVariables = strRSVariables & rsVal(0) & " "
						Case "datetime"
							strRSVariables = strRSVariables & "Convert(datetime,'" & SaveSqlDate(FormatDate(rsVal(0), False)) & "',120) "
						Case "nvarchar"
							strRSVariables = strRSVariables & "N'" & rsVal(0) & "' "
					End Select
				Case Else
					Select Case rsVar("varDataType")
						Case "nvarchar"
							strRSVariables = strRSVariables & "'' "
						Case "datetime"
							strRSVariables = strRSVariables & "'01/01/01' "
						Case "numeric", "int"
							strRSVariables = strRSVariables & "0 "
					End Select
			End Select
		rsVar.movenext
		loop
	End If
	getRSVariables = strRSVariables
End Function
Function getGenQry
	sqlGenQry = "select varVar, varDataType, varMaxChar, DefValBy, DefValDate, DefValValue, OLKCommon.dbo.DBOLKGetRSVarBaseIndex" & Session("ID") & "(rsIndex, varIndex) BaseIndex from OLKRSvars where rsIndex = " & Request("rsIndex")
	set rs = conn.execute(sqlGenQry)
	sqlGenQry = "declare @LanID int "
	do while not rs.eof
		If rs("varDataType") = "nvarchar" Then
			MaxChar = "(" & rs("varMaxChar") & ")"
		ElseIf rs("varDataType") = "numeric" Then
			MaxChar = "(19,6)"
		Else
			MaxChar = ""
		End If
		sqlGenQry = sqlGenQry & "declare @" & rs("varVar") & " " & rs("varDataType") & MaxChar & " set @" & rs("VarVar") & " = "
		Select Case rs("DefValBy")
			Case "V"
				Select Case rs("varDataType")
					Case "int", "numeric"
						sqlGenQry = sqlGenQry & rs("DefValValue") & " "
					Case "datetime"
						sqlGenQry = sqlGenQry & "Convert(datetime,'" & SaveSqlDate(FormatDate(rs("DefValDate"), False)) & "',120) "
					Case "nvarchar"
						sqlGenQry = sqlGenQry & "N'" & rs("DefValValue") & "' "
				End Select
			Case "Q"
				set rsVal = Server.CreateObject("ADODB.RecordSet")
				sqlVal = getRSVariables(rs("baseIndex")) & " " & rs("DefValValue")
				set rsVal = conn.execute(sqlVal)
				Select Case rs("varDataType")
					Case "int", "numeric"
						sqlGenQry = sqlGenQry & rsVal(0) & " "
					Case "datetime"
						sqlGenQry = sqlGenQry & "Convert(datetime,'" & SaveSqlDate(FormatDate(rsVal(0), False)) & "',120) "
					Case "nvarchar"
						sqlGenQry = sqlGenQry & "N'" & rsVal(0) & "' "
				End Select
			Case Else
				Select Case rs("varDataType")
					Case "nvarchar"
						sqlGenQry = sqlGenQry & "'' "
					Case "datetime"
						sqlGenQry = sqlGenQry & "'01/01/01' "
					Case "numeric", "int"
						sqlGenQry = sqlGenQry & "0 "
				End Select
		End Select
	rs.movenext
	loop
	
	Select Case UserType 
		Case "C"
			sqlGenQry = sqlGenQry & " declare @CardCode nvarchar(15) set @CardCode = '' "
		Case "V" 
		sqlGenQry = sqlGenQry & " declare @SlpCode int set @SlpCode = -1 "
	End Select
	
	sqlQuery = "select rsQuery, rsTop from OLKRS where rsIndex = " & Request("rsIndex")
	set rs = conn.execute(sqlQuery)
	sqlQuery = QueryFunctions(rs("rsQuery"))
	sqlGenQry = sqlGenQry & sqlQuery
	If rs("rsTop") = "Y" Then sqlGenQry = Replace(sqlGenQry, "@top", 1)
	getGenQry = sqlGenQry
End Function %>
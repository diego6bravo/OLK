<!--#include file="top.asp" -->
<!-- #INCLUDE file="FCKeditor/fckeditor.asp" -->
<!--#include file="lang/adminAnonLogin.asp" -->
<!--#include file="adminTradSubmit.asp"-->
<head>
<% 
conn.execute("use [" & Session("OLKDB") & "]")
set rd = Server.CreateObject("ADODB.recordset")
set rs = Server.CreateObject("ADODB.recordset")
If IsNull(myApp.AnTerms) or myApp.AnTerms = "" Then 
	AnTerms = "<div><font face=""Verdana"" size=""1"">&nbsp;</font></div>" 
Else 
	AnTerms= myApp.AnTerms
End If
%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<script language="javascript">
<% If Request("WebAddErr") = "True" Then %>
alert('<%=getadminAnonLoginLngStr("LtxtValAddress")%>"'.replace('{0}', '<%=Request("Address")%>').replace('{1}', '<%=Request("usedBy")%>'));
<% End If %>
function valFrm()
{
	<% If myApp.SVer >= 6 Then %>
	if (document.Form1.EnableAnSesion.checked && document.Form1.WebAddress.value == "")
	{
		alert('<%=getadminAnonLoginLngStr("LtxtValNoAddress")%>');
		document.Form1.WebAddress.focus();
		return false;
	}
	else if (document.Form1.EnableAnSesion.checked && document.Form1.AnSesListNum.selectedIndex <= 0)
	{
		alert('<%=getadminAnonLoginLngStr("LtxtValPList")%>');
		document.Form1.AnSesListNum.focus();
		return false;
	}
	<% End If %>
	/*
	else if ((document.Form1.AnRegAct.value == 'A' || document.Form1.AnRegAct.value == 'B') && !document.Form1.AnRegAct.disabled && !emailCheck(document.Form1.RegActMailAdd.value))
	{
		alert("<%=getadminAnonLoginLngStr("LtxtValRegEMail")%>");
		document.Form1.RegActMailAdd.focus();
		return false;
	}*/
	else if (document.Form1.AnonSesFilter.value != '' && document.Form1.valAnonSesFilter.value == 'Y')
	{
		alert("<%=getadminAnonLoginLngStr("LtxtValQryVal")%>");
		document.Form1.btnVerfyFilter.focus();
		return false;
	}
	return true;
}
function enableAnReg(chk)
{
	if (chk.checked)
	{
		document.Form1.AnRegAct.disabled=false;
		document.Form1.EnableAnRegTerms.disabled=false;
	}
	else
	{
		document.Form1.AnRegAct.disabled=true;
		document.Form1.EnableAnRegTerms.disabled=true;
	}
	enableAnRegAct(document.Form1.AnRegAct);
}
function enableAnRegAct(cmb)
{
	if ((cmb.value == 'A' || cmb.value == 'B' || cmb.value == 'C') && !cmb.disabled)
	{
		document.Form1.RegActMailAdd.readOnly=false;
		document.Form1.RegActMailAdd.style.backgroundColor="#D9F0FD";
	}
	else
	{
		document.Form1.RegActMailAdd.readOnly=true;
		document.Form1.RegActMailAdd.style.backgroundColor="#D4D0C8";
	}
	if ((cmb.value == 'C' || cmb.value == 'B') && !cmb.disabled)
	{
		document.Form1.AnRegConfAsignSLP.disabled = false;
		document.Form1.AnRegConfRejNote.disabled = false;
		document.Form1.AnRegConfFrom.readOnly = false;
		document.Form1.AnRegConfFrom.style.backgroundColor = '#D9F0FD';
		document.Form1.AnRegConfTo.readOnly = false;
		document.Form1.AnRegConfTo.style.backgroundColor = '#D9F0FD';
	}
	else
	{
		document.Form1.AnRegConfAsignSLP.disabled = true;
		document.Form1.AnRegConfRejNote.disabled = true;
		document.Form1.AnRegConfFrom.readOnly = true;
		document.Form1.AnRegConfFrom.style.backgroundColor = '#D4D0C8';
		document.Form1.AnRegConfTo.readOnly = true;
		document.Form1.AnRegConfTo.style.backgroundColor = '#D4D0C8';
	}
}

function Start(page, w, h, s) {
OpenWin = this.open(page, "DatePicker", "toolbar=no,menubar=no,location=no,scrollbars="+s+",resizable=no, width="+w+",height="+h);
}

function noteEdit(cmd) 
{
	var ReasonIndex = document.Form1.RejectNotes.value;
	var page = 'adminAcctRejReasons.asp?rIndex=' + ReasonIndex + '&cmd=' + cmd + '&pop=Y'
	
	OpenWin = this.open(page, "ReaesonEdit", "toolbar=no,menubar=no,location=no,scrollbars=no,resizable=no, width=364,height=310");
}
function chkThis(fld, min, oldVal)
{
	if (!IsNumeric(fld.value))
	{
		alert("<%=getadminAnonLoginLngStr("DtxtValNumVal")%>");
		fld.value = oldVal.value;
	}
	else if (parseFloat(fld.value) < parseFloat(min))
	{
		alert("<%=getadminAnonLoginLngStr("DtxtValNumMinVal")%>".replace('{0}', min));
		fld.value = min;
	}
	else if (parseFloat(fld.value) > 32727)
	{
		alert("<%=getadminAnonLoginLngStr("DtxtValNumMaxVal")%>".replace('{0}', 32727));
		fld.value = 32727;
	}
	fld.value = parseInt(fld.value);
	oldVal.value = fld.value;
}
</script>
</head>

<form method="POST" action="adminsubmit.asp" name="Form1" onsubmit="javascript:return valFrm()">
<%
strFormName = "Form1"
strTextAreaName = "anTerms"
%>
<table border="0" cellpadding="0" width="100%" id="table3">
	<tr>
		<td bgcolor="#E1F3FD"><b><font face="Verdana" size="2">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><%=getadminAnonLoginLngStr("LttlAnonSes")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1">
		<font color="#4783C5"><%=getadminAnonLoginLngStr("LttlAnonSesNote")%></font></font></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<div align="left">
			<table border="0" cellpadding="0" width="100%" id="table10">
				<tr>
					<td bgcolor="#F7FBFF" width="300">
					<img src="images/ganchito.gif"><font color="#4783C5"> </font>
					<input type="checkbox" class="noborder" name="EnableAnSesion" value="Y" <% If myApp.EnableAnSesion Then %>checked<%end if %> id="EnableAnSesion"><font face="Verdana" size="1" color="#4783C5"><label for="EnableAnSesion"><%=getadminAnonLoginLngStr("DtxtEnable")%></label></font></td>
					<td bgcolor="#F7FBFF">
					<p>
					&nbsp;</td>
				</tr>
				<tr>
					<td bgcolor="#F7FBFF" width="300">
					<img src="images/ganchito.gif"><font color="#4783C5">
					</font><font face="Verdana" size="1" color="#4783C5">
					<%=getadminAnonLoginLngStr("LtxtWebAddress")%></font></td>
					<td bgcolor="#F7FBFF">
					<p>
					<input type="text" name="WebAddress" size="45" class="input" dir="ltr" value="<%=myApp.WebAddress%>" onkeydown="return chkMax(event, this, 100);"></td>
				</tr>
				<tr>
					<td bgcolor="#F7FBFF" width="300">
					<img src="images/ganchito.gif"><font color="#4783C5"> </font>
					<font face="Verdana" size="1" color="#4783C5"><%=getadminAnonLoginLngStr("LtxtPList")%></font></td>
					<td bgcolor="#F7FBFF">
					<p>
					<font face="Verdana" size="1">
				<select size="1" name="AnSesListNum" class="input">
				<option></option>
			    <% 
			    set cmd = Server.CreateObject("ADODB.Command")
				cmd.ActiveConnection = connCommon
				cmd.CommandType = &H0004
				cmd.CommandText = "DBOLKGetPriceList" & Session("ID")
				cmd.Parameters.Refresh()
				cmd("@LanID") = Session("LanID")
				set rd = cmd.execute()
				do While NOT RD.EOF %>
				<option <% If rd("ListNum") = myApp.AnSesListNum Then Response.write "selected" %> value="<%=rd(0)%>"><%=myHTMLEncode(rd(1))%></option>
				<%  RD.MoveNext
				loop %>
				</select></font></td>
				</tr>
				</table>
		</div>
	</tr>
<tr>
		<td bgcolor="#E1F3FD">&nbsp;<b><font face="Verdana" size="1" color="#31659C"><%=getadminAnonLoginLngStr("LttlCatFilter")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE"><font face="Verdana" size="1">
		<img src="images/lentes.gif">
		<font color="#4783C5"><%=getadminAnonLoginLngStr("LttlCatFilterNote")%></font></font></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<div align="left">
			<table border="0" cellpadding="0" width="100%" id="table9">
				<tr>
					<td width="300" valign="top">
					<font face="Verdana" size="1" color="#4783C5"> 
					<%=getadminAnonLoginLngStr("DtxtQuery")%> - (ItemCode not in)</font></td>
					<td>
					<table cellpadding="0" cellspacing="0" border="0">
						<tr>
							<td rowspan="2">
								<textarea rows="10" dir="ltr" name="AnonSesFilter" cols="87" onkeydown="javascript:document.Form1.btnVerfyFilter.src='images/btnValidate.gif';document.Form1.btnVerfyFilter.style.cursor = 'hand';;document.Form1.valAnonSesFilter.value='Y';"><% If Not IsNull(myApp.AnonSesFilter) Then %><%=Server.HTMLEncode(myApp.AnonSesFilter)%><% End If %></textarea>
							</td>
							<td valign="top">
								<img style="cursor: hand; " src="images/qry_note.gif" id="btnNoteFilter" alt="<%=getadminAnonLoginLngStr("DtxtDefinition")%>" onclick="javascript:doFldNote(4, 'AnonSesFilter', -1, null);">
							</td>
						</tr>
						<tr>
							<td valign="bottom">
								<img src="images/btnValidateDis.gif" id="btnVerfyFilter" alt="<%=getadminAnonLoginLngStr("DtxtValidate")%>" onclick="javascript:if (document.Form1.valAnonSesFilter.value == 'Y')VerfyFilter();">
								<input type="hidden" name="valAnonSesFilter" value="N">	
						</td>
						</tr>
					</table>
					</td>
				</tr>

				</table>
		</div>
		</td>
	</tr>
	<tr>
		<td bgcolor="#E1F3FD"><b><font face="Verdana" size="2">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><%=getadminAnonLoginLngStr("LttlOptManRef")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1">
		<font color="#4783C5"><%=getadminAnonLoginLngStr("LttlOptManRefNote")%></font></font></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<div align="left">
			<table border="0" cellpadding="0" width="100%" id="table10">
				<% If 1 = 2 Then %>
				<tr>
					<td bgcolor="#F7FBFF">
					<img src="images/ganchito.gif"><font color="#4783C5">
					<font face="Verdana" size="1"><%=getadminAnonLoginLngStr("LtxtRemPwdAddr")%></font></font></td>
					<td bgcolor="#F7FBFF">
					<input type="text" dir="ltr" name="RemPwdMailAdd" size="45" class="input" value="<%=myApp.RemPwdMailAdd%>" onkeydown="return chkMax(event, this, 100);"></td>
					</tr>
				<% Else %>
				<input type="hidden" name="RemPwdMailAdd" value="<%=myApp.RemPwdMailAdd%>">
				<% End If %>
				<tr>
					<td bgcolor="#F7FBFF" width="425">
					<img src="images/ganchito.gif"><font color="#4783C5"> 
					<font face="Verdana" size="1"><%=getadminAnonLoginLngStr("LtxtClientType")%></font></font></td>
					<td bgcolor="#F7FBFF">
					<p>
					<select size="1" name="ClientType" class="input">
				    <option <% If myApp.ClientType = "C" Then %>selected<% End If %> value="C">
					<%=getadminAnonLoginLngStr("DtxtCmp")%></option>
					<option <% If myApp.ClientType = "P" Then %>selected<% End If %> value="P">
					<%=getadminAnonLoginLngStr("LtxtRegPerson")%></option>
				    </select></td>
				</tr>
				<% If myApp.LawsSet = "MX" or myApp.LawsSet = "CR" or myApp.LawsSet = "GT" or myApp.LawsSet = "US" or myApp.LawsSet = "CA" or myApp.LawsSet = "BR" Then %>
				<tr>
					<td bgcolor="#F7FBFF" width="425">
					<img src="images/ganchito.gif"><font color="#4783C5"> </font>
					<input class="noborder" type="checkbox" name="EnChooseCType" <% If myApp.EnChooseCType Then %>checked<% End If %> value="Y" id="EnChooseCType"><font face="Verdana" size="1" color="#4783C5"><label for="EnChooseCType"><%=getadminAnonLoginLngStr("LtxtEnChooseCType")%></label></font></td>
					<td bgcolor="#F7FBFF">
					<p>
					&nbsp;</td>
				</tr>
				<% End If %>
				<tr>
					<td bgcolor="#F7FBFF" width="425">
					<img src="images/ganchito.gif"><font color="#4783C5">
					</font>
					<input class="noborder" type="checkbox" name="EnableAnReg" value="Y" <% If myApp.EnableAnReg Then %>checked<%end if %> onclick="javascript:enableAnReg(this)" id="EnableAnReg"><font face="Verdana" size="1" color="#4783C5"><label for="EnableAnReg"><%=getadminAnonLoginLngStr("LtxtEnableAnReg")%></label></font></td>
					<td bgcolor="#F7FBFF">
					<p>
					&nbsp;</td>
				</tr>
				<tr>
					<td bgcolor="#F7FBFF" width="425">
					<img src="images/ganchito.gif"><font color="#4783C5">
					<font face="Verdana" size="1"><%=getadminAnonLoginLngStr("LtxtAnRegAct")%></font></font></td>
					<td bgcolor="#F7FBFF">
					<p>
					<select size="1" name="AnRegAct" class="input" <% If Not myApp.EnableAnReg Then %>disabled<% End If %> onchange="javascript:enableAnRegAct(this)">
					<option <% If myApp.AnRegAct = "N" Then %>selected<% End If %> value="N">
					<%=getadminAnonLoginLngStr("LtxtAnRegActAut")%></option>
					<option <% If myApp.AnRegAct = "A" Then %>selected<% End If %> value="A">
					<%=getadminAnonLoginLngStr("LtxtAnRegActEMail")%></option>
					<option <% If myApp.AnRegAct = "C" Then %>selected<% End If %> value="C">
					<%=getadminAnonLoginLngStr("LtxtAnRegActConf")%></option>
					<option <% If myApp.AnRegAct = "B" Then %>selected<% End If %> value="B">
					<%=getadminAnonLoginLngStr("LtxtAnRegActMailConf")%></option>
					</select></td>
				</tr>
				<% If 1 = 2 Then %>
				<tr>
					<td bgcolor="#F7FBFF" width="425">
					<img src="images/ganchito.gif"><font color="#4783C5">
					</font><font face="Verdana" size="1" color="#4783C5">
					<%=getadminAnonLoginLngStr("LtxtRegActMailAdd")%></font></td>
					<td bgcolor="#F7FBFF">
					<p>
					<input type="text" dir="ltr" <% If Not myApp.EnableAnReg or myApp.EnableAnReg and myApp.AnRegAct = "N" Then %>readonly<%end if %> name="RegActMailAdd" size="45" class="input" style="background-color: <% If myApp.AnRegAct = "N" or Trim(myApp.AnRegAct) = "" Then %>#D4D0C8<% ElseIf myApp.AnRegAct = "A" or myApp.AnRegAct = "B" or myApp.AnRegAct = "C" Then %>#D9F0FD<% End If %>" value="<%=myApp.RegActMailAdd%>" onkeydown="return chkMax(event, this, 100);"></td>
				</tr>
				<% Else %>
				<input type="hidden" name="RegActMailAdd" value="<%=myApp.RegActMailAdd%>">
				<% End If %>
				<tr>
					<td bgcolor="#F7FBFF" width="425">
					<img src="images/ganchito.gif"><font color="#4783C5">
					<font face="Verdana" size="1"><%=getadminAnonLoginLngStr("LtxtAnRegConfTime")%></font></font></td>
					<td bgcolor="#F7FBFF">
					<font face="Verdana" size="1" color="#4783C5"><%=getadminAnonLoginLngStr("DtxtFrom")%>
					<input <% If myApp.AnRegAct <> "C" and myApp.AnRegAct <> "B" Then %>readonly<% End If %> type="text" class="input" name="AnRegConfFrom" size="2" style="background-color: <% If myApp.AnRegAct <> "C" and myApp.AnRegAct <> "B" Then %>#D4D0C8<% Else %>#D9F0FD<% End If %>" value="<%=myApp.AnRegConfFrom%>" onchange="chkThis(this, 1, document.Form1.oldAnRegConfFrom)" MaxLength="2" onfocus="this.select()">
					<input type="hidden" id="oldAnRegConfFrom" name="oldAnRegConfFrom" value="<%=myApp.AnRegConfFrom%>">
					<font face="Verdana" size="1" color="#4783C5"><%=getadminAnonLoginLngStr("DtxtTo")%>
					<input <% If myApp.AnRegAct <> "C" and myApp.AnRegAct <> "B" Then %>readonly<% End If %> type="text" class="input" name="AnRegConfTo" size="2" style="background-color: <% If myApp.AnRegAct <> "C" and myApp.AnRegAct <> "B" Then %>#D4D0C8<% Else %>#D9F0FD<% End If %>" value="<%=myApp.AnRegConfTo%>" onchange="chkThis(this, 1, document.Form1.oldAnRegConfTo)" MaxLength="2" onfocus="this.select()"> 
					<input type="hidden" id="oldAnRegConfTo" name="oldAnRegConfTo" value="<%=myApp.AnRegConfTo%>">
					(<%=getadminAnonLoginLngStr("LtxtHours")%>)</font></td>
				</tr>
				<tr>
					<td bgcolor="#F7FBFF" width="425">
					<img src="images/ganchito.gif"><font color="#4783C5">
					</font>
					<input class="noborder" <% If myApp.AnRegAct <> "C" and myApp.AnRegAct <> "B" Then %>disabled<% End If %> type="checkbox" name="AnRegConfAsignSLP" value="Y" <% If myApp.AnRegConfAsignSLP Then %>checked<%end if %> id="AnRegConfAsignSLP"><font face="Verdana" size="1" color="#4783C5"><label for="AnRegConfAsignSLP"><%=getadminAnonLoginLngStr("LtxtAnRegConfAsignSLP")%></label></font></td>
					<td bgcolor="#F7FBFF">
					<p>
					&nbsp;</td>
				</tr>
				<tr>
					<td bgcolor="#F7FBFF" width="425">
					<img src="images/ganchito.gif">
					<input class="noborder" <% If myApp.AnRegAct <> "C" and myApp.AnRegAct <> "B" Then %>disabled<% End If %> type="checkbox" name="AnRegConfRejNote" value="Y" <% If myApp.AnRegConfRejNote Then %>checked<%end if %> id="AnRegConfRejNote"><font face="Verdana" size="1" color="#4783C5"><label for="AnRegConfRejNote"><%=getadminAnonLoginLngStr("LtxtAnRegConfRejNote")%></label></font></td>
					<td bgcolor="#F7FBFF">
					<p>
					&nbsp;</td>
				</tr>
				<tr>
					<td bgcolor="#F7FBFF" width="425">
					<img src="images/ganchito.gif"><font color="#4783C5">
					<font face="Verdana" size="1"><%=getadminAnonLoginLngStr("LtxtRejectNotes")%></font></font></td>
					<td bgcolor="#F7FBFF">
					<p>
					<select size="1" name="RejectNotes" class="input">
					<% sql1 = "select ReasonIndex, ReasonName from OLKAcctRejectNotes"
					set rd = conn.execute(sql1) %>
					<option value=""></option>
					<% do While NOT RD.EOF %>
					<option value="<%=rd("ReasonIndex")%>"><%=myHTMLEncode(rd("ReasonName"))%></option>
					<% rd.movenext
					loop %>
					</select> <a href="javascript:if(document.Form1.RejectNotes.selectedIndex>0)noteEdit('e')">
					<img height="16" src="images/wpedit.jpg" width="17" border="0" alt="<%=getadminAnonLoginLngStr("LtxtEditNote")%>"></a>
					<a href="javascript:noteEdit('a')">
					<img height="13" src="images/newdoc.gif" width="11" align="top" border="0" alt="<%=getadminAnonLoginLngStr("LtxtNewNote")%>"></a><a href="javascript:if(document.Form1.RejectNotes.selectedIndex>0)if(confirm('<%=getadminAnonLoginLngStr("LtxtConfDelReason")%>'))window.location.href='adminSubmit.asp?submitCmd=adminAcctRejReasons&cmd=d&rIndex=' + document.Form1.RejectNotes.value;"><img height="16" src="images/remove.gif" width="16" align="top" border="0" alt="<%=getadminAnonLoginLngStr("LtxtRemNote")%>"></a></td>
				</tr>
				<tr>
					<td bgcolor="#F7FBFF" width="425">
					<img src="images/ganchito.gif"><font color="#4783C5"> 
					</font>
					<input class="noborder" type="checkbox" name="EnableAnRegTerms" value="Y" <% If myApp.EnableAnRegTerms Then %>checked<%end if %> <% If Not myApp.EnableAnReg Then %>disabled<% End If %> id="EnableAnRegTerms"><font face="Verdana" size="1" color="#4783C5"><label for="EnableAnRegTerms"><%=getadminAnonLoginLngStr("LtxtEnableAnRegTerms")%></label></font></td>
					<td bgcolor="#F7FBFF">
					<p>
					&nbsp;</td>
				</tr>
				<tr>
					<td bgcolor="#F7FBFF" width="425">
					<img src="images/ganchito.gif">
					<font face="Verdana" size="1" color="#4783C5"><%=getadminAnonLoginLngStr("LtxtAnTerms")%></font></td>
					<td bgcolor="#F7FBFF">
					&nbsp;</td>
				</tr>
				<tr>
					<td bgcolor="#F7FBFF" colspan="2" width="800">
					<table cellpadding="0" cellspacing="0" border="0" width="100%">
					<tr>
						<td><%
									Dim oFCKeditor
									Set oFCKeditor = New FCKeditor
									oFCKeditor.BasePath = "FCKeditor/"
									oFCKeditor.Height = 300
									oFCKEditor.ToolbarSet = "Custom"
									If Not IsNull(AnTerms) Then oFCKEditor.Value = AnTerms
									oFCKEditor.Config("AutoDetectLanguage") = False
									If Session("myLng") <> "pt" Then
										oFCKEditor.Config("DefaultLanguage") = Session("myLng")
									Else
										oFCKEditor.Config("DefaultLanguage") = "pt-br"
									End If
									oFCKeditor.Create "AnTerms"
									%>
						</td>
						<td width="16" valign="bottom">
						<a href="javascript:doFldTrad('Common', '', '', 'AlterAnTerms', 'R', null);"><img src="images/trad.gif" alt="<%=getadminAnonLoginLngStr("DtxtTranslate")%>" border="0"></a>
						</td>
					</tr>
				</table>
				</td>
				</tr>
				</table>
		</div>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table5">
			<tr>
				<td width="77">
				<input type="submit" value="<%=getadminAnonLoginLngStr("DtxtSave")%>" name="B1" class="OlkBtn"></td>
				<td><hr color="#0D85C6" size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
</table>
<input type="hidden" name="submitCmd" value="adminAnonLogin">
</form>
<script language="javascript">
function VerfyFilter()
{
	document.frmVerfyQuery.Query.value = document.Form1.AnonSesFilter.value;
	if (document.frmVerfyQuery.Query.value != '')
	{
		document.frmVerfyQuery.submit();
	}
	else
	{
		VerfyQueryVerified();
	}
}
function VerfyQueryVerified()
{
	//document.Form1.btnVerfyFilter.disabled = true;
	document.Form1.btnVerfyFilter.src='images/btnValidateDis.gif'
	document.Form1.btnVerfyFilter.style.cursor = '';
	document.Form1.valAnonSesFilter.value='N';
}
//-->
</script>
<form name="frmVerfyQuery" action="verfyQuery.asp" method="post" target="iVerfyQuery">
	<iframe id="iVerfyQuery" name="iVerfyQuery" style="display: none" src=""></iframe>
	<input type="hidden" name="type" value="AnonCatFilter">
	<input type="hidden" name="Query" value="">
	<input type="hidden" name="parent" value="Y">
</form>
<!--#include file="bottom.asp" -->
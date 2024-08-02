<% pageTtl = "Agregar Reporte"
pDesc = pageTtl %>
<!--#include file="repTop.inc" -->
<!--#include file="lang/adminRepNew.asp" -->
<!--#include file="adminTradSubmit.asp"-->



<%
SQL = "Select rgIndex, rgName from " & repTbl & "RG where UserType = '" & Request("UserType") & "' and rgIndex >= 0 order by rgName asc"
SET RS = Conn.Execute(SQL)
%>
<script language="javascript">
function valFrm()
{
	if (document.form1.rsName.value == '')
	{
		alert('<%=getadminRepNewLngStr("LtxtValRepNam")%>');
		document.form1.rsName.focus();
		return false;
	}
	<% If 1 = 2 Then %>
	else if (document.form1.rsQuery.value == '')
	{
		alert('<%=getadminRepNewLngStr("LtxtValRepQry")%>');
		document.form1.rsQuery.focus();
		return false;
	}
	else if (document.form1.valRSQuery.value == 'Y')
	{
		alert("<%=getadminRepNewLngStr("LtxtValRepQryVal")%>");
		document.form1.btnVerfy.focus();
		return false;
	}
	<% End If %>
	return true;
}
<% If 1 = 2 Then %>
function VerfyQuery()
{
	document.frmVerfyQuery.Query.value = document.form1.rsQuery.value;
	document.frmVerfyQuery.submit();
}

function VerfyQueryVerified()
{
	//document.form1.btnVerfy.disabled = true;

	document.form1.btnVerfy.src='images/btnValidateDis.gif'
	document.form1.btnVerfy.style.cursor = '';
	document.form1.valRSQuery.value='N';
}
<% End If %>
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
</script>
<form name="frmVerfyQuery" action="verfyQuery.asp" method="post" target="iVerfyQuery">
	<input type="hidden" name="type" value="newQuery">
	<input type="hidden" name="Query" value="">
	<input type="hidden" name="AddPath" value="">
	<input type="hidden" name="UserType" value="<%=Request("UserType")%>">
	<input type="hidden" name="parent" value="Y">
</form>
<form name="form1" action="repSubmit.asp" method="POST" onsubmit="javascript:return valFrm();">
<input type="hidden" name="rsNameTrad">
<input type="hidden" name="rsDescTrad">
<input type="hidden" name="rsQueryDef">
<table border="0" cellpadding="0" width="98%" id="table3">
	<tr class="TblRepTlt">
		<td>&nbsp;<%=getadminRepNewLngStr("LttlAddRep")%></td>
	</tr>
	<tr class="TblRepNrm">
		<td>
		<p align="justify"><img src="images/lentes.gif">
		<%=getadminRepNewLngStr("LttlAddRepNote")%></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table7">
			<tr class="TblRepNrm">
				<td width="87">
				<%=getadminRepNewLngStr("DtxtName")%></td>
				<td>
    <input type="text" name="rsName" size="67" value="" onkeydown="return chkMax(event, this, 60);"><a href="javascript:doFldTrad('RS', 'rsIndex', '', 'alterRSName', 'T', form1.rsNameTrad);"><img src="images/trad.gif" alt="|D:txtTranslate|" border="0"></a></td>
			</tr>
			<tr class="TblRepNrm">
				<td width="87" valign="middle">
				<%=getadminRepNewLngStr("DtxtGroup")%></td>
				<td><select name="rgIndex" size="1">
				<% do while not rs.eof %>
				<option value="<%=rs("rgIndex")%>"><%=myHTMLEncode(rs("rgName"))%></option>
				<% rs.movenext
				loop
				set rs = nothing %>
				</select>&nbsp;</td>
			</tr>
			<tr class="TblRepNrm">
				<td width="76">
				<%=getadminRepNewLngStr("LtxtTime")%></td>
				<td width="476">
				<select size="1" name="cmbRefresh">
				<option value="0"><%=getadminRepNewLngStr("DtxtDisabled")%></option>
				<option <% If SelectedValue = 1 Then %>selected<% End If %> value="1"><%=getadminRepNewLngStr("DtxtMinute")%></option>
				<option <% If SelectedValue = 5 Then %>selected<% End If %> value="5">5 <%=getadminRepNewLngStr("DtxtMinutes")%></option>
				<option <% If SelectedValue = 10 Then %>selected<% End If %> value="10">10 <%=getadminRepNewLngStr("DtxtMinutes")%></option>
				<option <% If SelectedValue = 30 Then %>selected<% End If %> value="30">30 <%=getadminRepNewLngStr("DtxtMinutes")%></option>
				<option <% If SelectedValue = 60 Then %>selected<% End If %> value="60"><%=getadminRepNewLngStr("DtxtHour")%></option>
				</select></td>
				</tr>
			<tr class="TblRepNrm">
				<td width="87" valign="top">
				<%=getadminRepNewLngStr("DtxtDescription")%></td>
				<td>
				<table cellpadding="0" cellspacing="0" border="0" width="100%">
					<tr>
						<td><textarea rows="3" name="rsDesc" cols="88" style="width: 100%"></textarea>
						</td>
						<td width="24" valign="bottom"><a href="javascript:doFldTrad('RS', 'rsIndex', '', 'alterRSDesc', 'M', document.form1.rsDescTrad);"><img src="images/trad.gif" alt="|D:txtTranslate|" border="0"></a></td>
					</tr>
				</table>
			    </td>
			</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table5">
			<tr>
				<td width="77">
				<input class="BtnRep" type="submit" value="<%=getadminRepNewLngStr("DtxtAdd")%>" name="B1"></td>
				<td><hr size="1"></td>
				<td width="77">
				<input type="button" class="BtnRep" value="<%=getadminRepNewLngStr("DtxtCancel")%>" name="btnCancel" onclick="javascript:if(confirm('<%=getadminRepNewLngStr("DtxtConfCancel")%>'))window.location.href='adminReps.asp?uType=<%=Request("UserType")%>'"></td>
			</tr>
		</table>
		</td>
	</tr>
</table>
<input type="hidden" name="cmd" value="newRep">
<iframe id="iVerfyQuery" name="iVerfyQuery" style="display: none" src=""></iframe>
</form>
<!--#include file="repBottom.inc" -->
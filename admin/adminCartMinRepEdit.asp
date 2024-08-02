<!-- #include file="top.asp" -->
<!--#include file="lang/adminCartMinRepEdit.asp" -->
<!--#include file="adminTradSubmit.asp"-->
<!--#include file="adminTradSave.asp"-->
<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style type="text/css">
.style1 {
	background-color: #F3FBFE;
}
.style4 {
				font-weight: normal;
				color: #4783C5;
}
.style5 {
				font-family: Verdana;
}
.style6 {
				background-color: #F3FBFE;
				font-family: Verdana;
				font-size: xx-small;
				color: #4783C5;
}
.style7 {
				color: #4783C5;
}
</style>
</head>
<script language="javascript" src="js_up_down.js"></script>

<% conn.execute("use [" & Session("olkdb") & "]")
If Request("goAction") = "editLine" and Request("btnApply") = "" and Request("btnSave") = "" Then
	editIndex = Request("LineIndex")
	rowType = Request("RowType")
End If
If Request("goAction") = "updateLine" Then
	If Request("RowActive") = "Y" Then RowActive = "Y" Else RowActive = "N"
	If Request("RowAlign") <> "" Then Align = "'" & Request("RowAlign") & "'" Else Align = "NULL"
	If Request("ShowV") = "Y" Then ShowV = "Y" Else ShowV = "N"
	If Request("ShowC") = "Y" Then ShowC = "Y" Else ShowC = "N"
	If Request("PrintV") = "Y" Then PrintV = "Y" Else PrintV = "N"
	If Request("PrintC") = "Y" Then PrintC = "Y" Else PrintC = "N"
	If Request("RowQuery") <> "" Then RowQuery = "N'" & saveHTMLDecode(Request("RowQuery"), False) & "'" Else RowQuery = "NULL"
	If Request("SystemQuery") <> "" Then SystemQuery = "N'" & saveHTMLDecode(Request("SystemQuery"), False) & "'" Else SystemQuery = "NULL"
	If Request("btnRestore") <> "" Then RowQuery = "(select Convert(nvarchar(4000),RowQuery) from OLKCommon..OLKCMRep where RowType = 'S' and LineIndex = " & Request("LineIndex") & ") "
	sql = "update OLKCMREP set RowName = N'" & saveHTMLDecode(Request("RowName"), False) & "', RowActive = '" & RowActive & "', " & _
	"RowQuery = " & RowQuery & ", SystemQuery = " & SystemQuery & ", Align = " & Align & ", " & _
	"ShowV = '" & ShowV & "', ShowC = '" & ShowC & "', PrintV = '" & PrintV & "', PrintC = '" & PrintC & "', rowOrder = " & Request("RowOrder") & " " & _
	"where RowType = '" & Request("RowType") & "' and LineIndex = " & Request("LineIndex")
	conn.execute(sql)
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKGenQry" & Session("ID")
	cmd.Parameters.Refresh
	cmd("@Type") = "CMREP"
	cmd.execute()
	If Request("btnApply") <> "" or Request("btnRestore") <> "" Then 
		editIndex = Request("LineIndex")
		rowType = Request("RowType")
		Select Case rowType
			Case "S"
				DefID = 1
			Case "U"
				DefID = 2
		End Select
	End If
ElseIf Request("goAction") = "addLine" Then
	If Request("RowActive") = "Y" Then RowActive = "Y" Else RowActive = "N"
	If Request("RowAlign") <> "" Then Align = "'" & Request("RowAlign") & "'" Else Align = "NULL"
	If Request("ShowV") = "Y" Then ShowV = "Y" Else ShowV = "N"
	If Request("ShowC") = "Y" Then ShowC = "Y" Else ShowC = "N"
	If Request("PrintV") = "Y" Then PrintV = "Y" Else PrintV = "N"
	If Request("PrintC") = "Y" Then PrintC = "Y" Else PrintC = "N"
	If Request("RowQuery") <> "" Then RowQuery = "N'" & saveHTMLDecode(Request("RowQuery"), False) & "'" Else RowQuery = "NULL"
	If Request("SystemQuery") <> "" Then SystemQuery = "N'" & saveHTMLDecode(Request("SystemQuery"), False) & "'" Else SystemQuery = "NULL"
	sql = "declare @LineIndex int set @LineIndex = IsNull((select Max(LineIndex)+1 from OLKCMREP where RowType = 'U'),0) select @LineIndex LineIndex " & _
	"insert OLKCMREP(RowType, LineIndex, RowName, RowQuery, SystemQuery, RowActive, Align, ShowV, ShowC, PrintV, PrintC, RowOrder) " & _
	"values('U', @LineIndex, N'" & saveHTMLDecode(Request("RowName"), False) & "', " & RowQuery & ", " & SystemQuery & ", '" & RowActive & "', " & Align & ", '" & ShowV & "', '" & ShowC & "', '" & PrintV & "', '" & PrintC & "', " & Request("RowOrder") & ") "
	set rs = conn.execute(sql)
	If Request("btnApply") <> "" Then 
		editIndex = rs(0)
		rowType = "U"
		DefID = 2
	End If
	rs.close
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKGenQry" & Session("ID")
	cmd.Parameters.Refresh
	cmd("@Type") = "CMREP"
	cmd.execute()
	
	If Request("rowNameTrad") <> "" Then
		SaveNewTrad Request("rowNameTrad"), "CMREP", "RowType,LineIndex", "alterRowName", rowType & "," & editIndex
	End If
	
	If Request("RowQueryDef") <> "" Then
		SaveNewDef Request("RowQueryDef"), CStr(DefID) & CStr(editIndex)
	End If
	
	If Request("SystemQueryDef") <> "" Then
		SaveNewDef Request("SystemQueryDef"), CStr(DefID) & CStr(editIndex)
	End If

End If
If Request("btnSave") <> "" Then Response.Redirect "adminCartMinRep.asp" %>
<table border="0" cellpadding="0" width="100%" id="table3">
  	<script language="javascript">
  	function valFrm2()
  	{
  		if (document.form2.rowName.value == '') {
  			alert('<%=getadminCartMinRepEditLngStr("LtxtValFldNam")%>');
  			document.form2.rowName.focus();
  			return false; }
  		else if (document.form2.RowQuery.value == '' && document.form2.SystemQuery.value == '') {
  			alert('<%=getadminCartMinRepEditLngStr("LtxtValQry")%>');
  			document.form2.RowQuery.focus();
  			return false; }
  		else if (document.form2.valRowQuery.value == 'Y') {
  			alert('<%=getadminCartMinRepEditLngStr("LtxtValQryVal")%>');
  			document.form2.btnVerfy.focus();
  			return false; }
  		else if (document.form2.valSystemQuery.value == 'Y') {
  			alert('<%=getadminCartMinRepEditLngStr("LtxtValQryVal")%>');
  			document.form2.btnVerfySystem.focus();
  			return false; }
  		return true;
  	}
  	</script>
	<form method="POST" action="adminCartMinRepEdit.asp" name="form2" onsubmit="return valFrm2()">
	<tr>
		<td bgcolor="#E1F3FD">&nbsp;<b><font size="1" face="Verdana" color="#31659C"><% If Request("goAction") = "editLine" or Request("btnApply") <> "" or Request("btnRestore") <> "" Then %><%=getadminCartMinRepEditLngStr("LttlEditMinRepFld")%><% Else %><%=getadminCartMinRepEditLngStr("LttlAddMinRepFld")%><% End If %></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1">
		<font color="#4783C5"><%=getadminCartMinRepEditLngStr("LttlMinRepFldNote")%></font></font></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
		<% RowActive = ""
		If Request("goAction") = "editLine" or Request("btnApply") <> "" or Request("btnRestore") <> "" Then
			sql = "select * from OLKCMREP where RowType = '" & rowType & "' and LineIndex = " & editIndex
			set rs = conn.execute(sql) 
			rowName = rs("rowName")
			RowActive = rs("RowActive")
			RowQuery = rs("RowQuery")
			SystemQuery = rs("SystemQuery")
			Align = rs("Align")
			ShowV = rs("ShowV")
			ShowC = rs("ShowC")
			PrintV = rs("PrintV")
			PrintC = rs("PrintC")
			rowOrder = rs("rowOrder")
			Select Case rs("rowType")
				Case "S"
					DefID = 1
					rowType = "S"
				Case "U"
					DefID = 2
					rowType = "U"
			End Select
		Else
			RowActive = "N"
			ShowV = "N"
			ShowC= "N"
			PrintV = "N"
			PrintC = "N"
			rowName = ""
			rowType = "U"
			DefID = 2
			sql = "select IsNull(Max(rowOrder)+1, 0) from OLKCMREP"
			set rs = conn.execute(sql)
			rowOrder = rs(0) %>
		<input type="hidden" name="rowNameTrad">
		<input type="hidden" name="RowQueryDef">
		<input type="hidden" name="SystemQueryDef">
		<% End If %>
			<tr>
				<td bgcolor="#E2F3FC" style="width: 120px"><b>
				<font size="1" face="Verdana" color="#31659C"><%=getadminCartMinRepEditLngStr("DtxtName")%></font></b></td>
				<td valign="top" class="style1">
				<table cellpadding="0" cellspacing="0" border="0" width="300">
					<tr>
						<td>
						<input name="rowName" style="width: 100%; " class="input" value="<%=Server.HTMLEncode(rowName)%>" size="20" onkeydown="return chkMax(event, this, 50);">
						</td>
						<td width="16"><a href="javascript:doFldTrad('CMREP', 'RowType,LineIndex', '<%=Request("RowType")%>,<%=Request("LineIndex")%>', 'AlterRowName', 'T', <% If Request("goAction") = "editLine" or Request("btnApply") <> "" or Request("btnRestore") <> "" Then %>document.form2.rowNameTrad<% Else %>null<% End If %>);"><img src="images/trad.gif" alt="<%=getadminCartMinRepEditLngStr("DtxtTranslate")%>" border="0"></a></td>
					</tr>
				</table>
				</td>				
			</tr>
			<tr>
				<td bgcolor="#E2F3FC" style="width: 120px"><b>
				<font face="Verdana" size="1" color="#31659C"><%=getadminCartMinRepEditLngStr("DtxtOrder")%></font></b></td>
				<td valign="top" class="style1">
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td>
							<input type="text" name="RowOrder" id="RowOrder" size="7" style="font-size: 10px; font-family: Verdana; color: #3F7B96; font-weight: bold; border: 1px solid #68A6C0; background-color: #D9F0FD; text-align:right" class="input" onfocus="this.select();" onmouseup="event.preventDefault()" onkeydown="return chkMax(event, this, 6);" value="<%=rowOrder%>">
						</td>
						<td valign="middle">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td><img src="images/img_nud_up.gif" id="btnRowOrderUp"></td>
							</tr>
							<tr>
								<td><img src="images/spacer.gif"></td>
							</tr>
							<tr>
								<td><img src="images/img_nud_down.gif" id="btnRowOrderDown"></td>
							</tr>
						</table></td>
					</tr>
				</table></td>				
			</tr>
			<tr>
				<td bgcolor="#E2F3FC" style="width: 120px"><b>
				<font face="Verdana" size="1" color="#31659C"><%=getadminCartMinRepEditLngStr("DtxtAlignment")%></font></b></td>
				<td valign="top" class="style1">
				<select size="1" name="RowAlign">
				<option></option>
				<option <% If Align = "L" Then %>selected<% End If %> value="L"><%=getadminCartMinRepEditLngStr("DtxtLeft")%></option>
				<option <% If Align = "C" Then %>selected<% End If %> value="C"><%=getadminCartMinRepEditLngStr("DtxtCenter")%></option>
				<option <% If Align = "R" Then %>selected<% End If %> value="R"><%=getadminCartMinRepEditLngStr("DtxtRight")%></option>
				</select></td>				
			</tr>
			<tr>
				<td bgcolor="#E2F3FC" style="width: 120px">&nbsp;</td>
				<td valign="top" class="style6">
				<font face="Verdana" size="1" color="#31659C">
				<span class="style5"><font size="1"><span class="style7">
				<input type="checkbox" name="RowActive" class="noborder" id="RowActive" <% If RowActive = "Y" Then %>checked<% End If %> value="Y"></span></font></span></font><font face="Verdana" size="1"><span class="style4"><label for="RowActive"><%=getadminCartMinRepEditLngStr("DtxtActive")%></label></span></font></td>				
			</tr>
			<tr>
				<td bgcolor="#E2F3FC" style="width: 120px">&nbsp;</td>
				<td valign="top" class="style6">
				<font face="Verdana" size="1" color="#31659C">
				<span class="style5"><font size="1"><span class="style7">
				<input type="checkbox" name="ShowV" class="noborder" id="ShowV" <% If ShowV = "Y" Then %>checked<% End If %> value="Y"></span></font></span></font><font face="Verdana" size="1"><span class="style4"><label for="ShowV"><%=myHTMLDecode(getadminCartMinRepEditLngStr("LtxtAgentsVisible"))%></label></span></font></td>				
			</tr>
			<tr>
				<td bgcolor="#E2F3FC" style="width: 120px">&nbsp;</td>
				<td valign="top" class="style6">
				<font face="Verdana" size="1" color="#31659C">
				<span class="style5"><font size="1"><span class="style7">
				<input type="checkbox" name="PrintV" class="noborder" id="PrintV" <% If PrintV = "Y" Then %>checked<% End If %> value="Y"></span></font></span></font><font face="Verdana" size="1"><span class="style4"><label for="PrintV"><%=myHTMLDecode(getadminCartMinRepEditLngStr("LtxtAgentsPrinting"))%></label></span></font></td>				
			</tr>
			<tr>
				<td bgcolor="#E2F3FC" style="width: 120px">&nbsp;</td>
				<td valign="top" class="style6">
				<font face="Verdana" size="1" color="#31659C">
				<span class="style5"><font size="1"><span class="style7">
				<input type="checkbox" name="ShowC" class="noborder" id="ShowC" <% If ShowC = "Y" Then %>checked<% End If %> value="Y"></span></font></span></font><font face="Verdana" size="1"><span class="style4"><label for="ShowC"><%=myHTMLDecode(getadminCartMinRepEditLngStr("LtxtClientsVisible"))%></label></span></font></td>				
			</tr>
			<tr>
				<td bgcolor="#E2F3FC" style="width: 120px">&nbsp;</td>
				<td valign="top" class="style6">
				<font face="Verdana" size="1" color="#31659C">
				<span class="style5"><font size="1"><span class="style7">
				<input type="checkbox" name="PrintC" class="noborder" id="PrintC" <% If PrintC = "Y" Then %>checked<% End If %> value="Y"></span></font></span></font><font face="Verdana" size="1"><span class="style4"><label for="PrintC"><%=myHTMLDecode(getadminCartMinRepEditLngStr("LtxtClientsPrinting"))%></label></span></font></td>				
			</tr>
			<tr>
				<td bgcolor="#E2F3FC" style="width: 120px" valign="top">
						<b><font face="Verdana" size="1" color="#31659C">
						<%=getadminCartMinRepEditLngStr("DtxtQuery")%> (<%=getadminCartMinRepEditLngStr("DtxtOLK")%>)</font></b></td>	
						<td valign="top" class="style1">
				
						<table cellpadding="0" cellspacing="0" border="0" width="100%">
							<tr>
								<td rowspan="2">
									<textarea cols="78" dir="ltr" name="RowQuery" class="input" style="width: 100%; " rows="6" onkeypress="javascript:document.form2.btnVerfy.src='images/btnValidate.gif';document.form2.btnVerfy.style.cursor = 'hand';document.form2.valRowQuery.value='Y';"><%=myHTMLEncode(RowQuery)%></textarea>
								</td>
								<td valign="top" width="1">
									<img style="cursor: hand; " src="images/qry_note.gif" id="btnNoteFilter" alt="<%=getadminCartMinRepEditLngStr("DtxtDefinition")%>" onclick="javascript:doFldNote(16, 'RowQuery', '<%=DefID%><%=editIndex%>', <% If Request("goAction") = "editLine" or Request("btnApply") <> "" or Request("btnRestore") <> "" Then %>null<% Else %>document.form2.RowQueryDef<% End If %>);">
								</td>
							</tr>
							<tr>
								<td valign="bottom" width="1">
									<img src="images/btnValidateDis.gif" id="btnVerfy" alt="<%=getadminCartMinRepEditLngStr("DtxtValidate")%>" onclick="javascript:if (document.form2.valRowQuery.value == 'Y')VerfyQuery();">
									<input type="hidden" name="valRowQuery" value="N">
								</td>
							</tr>
						</table></td>					
			</tr>
			<tr>
				<td bgcolor="#E2F3FC" style="width: 120px" valign="top">
						<b><font face="Verdana" size="1" color="#31659C">
						<%=getadminCartMinRepEditLngStr("DtxtVariables")%></font></b></td>
				<td valign="top" class="style1">
						<font face="Verdana" size="1" color="#4783C5">
						<span dir="ltr">@LogNum</span> = 
						<%=getadminCartMinRepEditLngStr("LtxtLogNumDesc")%><br>
						<span dir="ltr">@CardCode</span> = <%=getadminCartMinRepEditLngStr("LtxtCCodeDesc")%><br>
						<span dir="ltr">@LanID</span> = <%=getadminCartMinRepEditLngStr("DtxtLanID")%></font></td>		
			</tr>
			<tr>
				<td bgcolor="#E2F3FC" style="width: 120px" valign="top">
						<b><font face="Verdana" size="1" color="#31659C">
						<%=getadminCartMinRepEditLngStr("DtxtQuery")%> (<%=getadminCartMinRepEditLngStr("DtxtSystem")%>)</font></b></td>	
						<td valign="top" class="style1">
				
						<table cellpadding="0" cellspacing="0" border="0" width="100%">
							<tr>
								<td rowspan="2">
									<textarea cols="78" dir="ltr" name="SystemQuery" class="input" style="width: 100%; " rows="6" onkeypress="javascript:document.form2.btnVerfySystem.src='images/btnValidate.gif';document.form2.btnVerfySystem.style.cursor = 'hand';document.form2.valSystemQuery.value='Y';"><%=myHTMLEncode(SystemQuery)%></textarea>
								</td>
								<td valign="top" width="1">
									<img style="cursor: hand; " src="images/qry_note.gif" id="btnNoteFilterSystem" alt="<%=getadminCartMinRepEditLngStr("DtxtDefinition")%>" onclick="javascript:doFldNote(16, 'SystemQuery', '<%=DefID%><%=editIndex%>', <% If Request("goAction") = "editLine" or Request("btnApply") <> "" or Request("btnRestore") <> "" Then %>null<% Else %>document.form2.SystemQueryDef<% End If %>);">
								</td>
							</tr>
							<tr>
								<td valign="bottom" width="1">
									<img src="images/btnValidateDis.gif" id="btnVerfySystem" alt="<%=getadminCartMinRepEditLngStr("DtxtValidate")%>" onclick="javascript:if (document.form2.valSystemQuery.value == 'Y')VerfyQuerySystem();">
									<input type="hidden" name="valSystemQuery" value="N">
								</td>
							</tr>
						</table></td>					
			</tr>
			<tr>
				<td bgcolor="#E2F3FC" style="width: 120px" valign="top">
						<b><font face="Verdana" size="1" color="#31659C">
						<%=getadminCartMinRepEditLngStr("DtxtVariables")%></font></b></td>
				<td valign="top" class="style1">
						<font face="Verdana" size="1" color="#4783C5">
						<span dir="ltr">{Table}</span> = <%=getadminCartMinRepEditLngStr("LtxtTableDesc")%>&nbsp;({Table} = INV, RDR, ...)<br><span dir="ltr">@DocEntry</span> = 
						<%=getadminCartMinRepEditLngStr("LtxtDocEntry")%><br>
						<span dir="ltr">@CardCode</span> = <%=getadminCartMinRepEditLngStr("LtxtCCodeDesc")%><br>
						<span dir="ltr">@LanID</span> = <%=getadminCartMinRepEditLngStr("DtxtLanID")%></font></td>		
			</tr>
			<tr>
				<td bgcolor="#E2F3FC" style="width: 120px">
						<b><font face="Verdana" size="1" color="#31659C">
						<%=getadminCartMinRepEditLngStr("DtxtFunctions")%></font></b></td>
				<td valign="top" class="style1">
				
						<% HideFunctionTitle = True
						functionClass="TblFlowFunction" %><!--#include file="myFunctions.asp"--></td>				
			</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table5">
			<tr>
				<td width="77">
				<input type="submit" value="<%=getadminCartMinRepEditLngStr("DtxtApply")%>" name="btnApply" class="OlkBtn"></td>
				<td width="77">
				<input type="submit" value="<%=getadminCartMinRepEditLngStr("DtxtSave")%>" name="btnSave" class="OlkBtn"></td>
				<td><hr color="#0D85C6" size="1"></td>
				<% If Request("RowType") = "S" Then %>
				<td width="77">
				<input type="submit" value="<%=getadminCartMinRepEditLngStr("DtxtRestore")%>" name="btnRestore" class="OlkBtn" onclick="javascript:return confirm('<%=getadminCartMinRepEditLngStr("LtxtValRestoreFld")%>');"></td><% End If %>
				<td width="77">
				<input type="button" value="<%=getadminCartMinRepEditLngStr("DtxtCancel")%>" name="btnCancel" class="OlkBtn" onclick="javascript:if(confirm('<%=getadminCartMinRepEditLngStr("LtxtValCanFld")%>'))window.location.href='adminCartMinRep.asp'"></td>
			</tr>
		</table>
		</td>
	</tr>
	<input type="hidden" name="LineIndex" value="<%=editIndex%>">
	<input type="hidden" name="goAction" value="<% If Request("goAction") = "editLine" or Request("btnApply") <> "" Then %>updateLine<% Else %>addLine<% End If %>">
	<input type="hidden" name="RowType" value="<%=Request("RowType")%>">
</form>
	<tr>
		<td height="15"></td>
	</tr>
</table>
<script language="javascript">
NumUDAttach('form2', 'RowOrder', 'btnRowOrderUp', 'btnRowOrderDown');
function VerfyQuery()
{
	document.frmVerfyQuery.type.value = 'minrep';
	document.frmVerfyQuery.Query.value = document.form2.RowQuery.value;
	document.frmVerfyQuery.submit();
}
function VerfyQuerySystem()
{
	document.frmVerfyQuery.type.value = 'minrepSys';
	document.frmVerfyQuery.Query.value = document.form2.SystemQuery.value;
	document.frmVerfyQuery.submit();
}

function VerfyQueryVerified()
{
	switch (document.frmVerfyQuery.type.value)
	{
		case 'minrep':
			document.form2.btnVerfy.src='images/btnValidateDis.gif'
			document.form2.btnVerfy.cursor = '';
			document.form2.valRowQuery.value='N';
			break;
		case 'minrepSys':
			document.form2.btnVerfySystem.src='images/btnValidateDis.gif'
			document.form2.btnVerfySystem.cursor = '';
			document.form2.valSystemQuery.value='N';
			break;
	}

	//document.form2.btnVerfy.disabled = true;
}
</script>
<form name="frmVerfyQuery" action="verfyQuery.asp" method="post" target="iVerfyQuery">
	<iframe id="iVerfyQuery" name="iVerfyQuery" style="display: none" src=""></iframe>
	<input type="hidden" name="type" value="">
	<input type="hidden" name="Query" value="">
	<input type="hidden" name="parent" value="Y">
</form>

<!-- #include file="bottom.asp" -->
<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="../chkLogin.asp" -->
<!--#include file="lang/viewRepVals.asp" -->
<!--#include file="../myHTMLEncode.asp"-->
<html <% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>>
<!--#include file="viewRepValsTop.inc" -->
<% 
Dim selectName
If Request("isSubmit") <> "Y" Then %>
<head>

<meta http-equiv="Content-Type" content="text/html; charset=utf-8" >
<% set rs = server.createobject("ADODB.RecordSet")
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetRSVars" & Session("ID")
cmd.Parameters.Refresh()
cmd("@rsIndex") = Request("rsIndex")
cmd("@LanID") = Session("LanID")
set rs = cmd.execute()
set rd = server.createobject("ADODB.RecordSet")
rsTop = rs("rsTop") = "Y"
rsTopDef = rs("rsTopDef")
%>
<title><%=rs("rsName")%></title>
<style>
<!--
input        { font-family: Verdana; font-size: 10px }
select       { font-family: Verdana; font-size: 10px }
label   	 { font-family: Verdana; font-size: 10px }

.noborder {
	border-style : solid;
	border-width : 0;
}
-->
</style>
</head>
<script language="javascript" src="../general.js"></script>
<script language="javascript">
var OpenWin = null;
var Field
function chkWin() { if (OpenWin != null) if (!OpenWin.closed) OpenWin.focus() }

function Start(o, page, w, h, s, r) {
Field = o
OpenWin = this.open(page, "queryWin", "toolbar=no,menubar=no,location=no,scrollbars="+s+",resizable="+r+", width="+w+",height="+h);
OpenWin.focus()
}

function setTimeStamp(Nothing, var1) {
	Field.value = var1;
	if (Field.onchange != null) Field.onchange();
}

function chkNum(fld, dType)
{
	if (dType != 'nvarchar')
	{
		if (!MyIsNumeric(fld.value))
		{
			alert('<%=getviewRepValsLngStr("DtxtValNumVal")%>');
			fld.value = '';
			fld.focus();
		}
		else if (dType == 'int')
		{
			fld.value = parseInt(fld.value);
		}
	}
}
var noVal = false;
function reload(targetIndex)
{
	noVal = true;
	if (targetIndex != '')
	{
		var arrIndex = targetIndex.toString().split(', ');
		for (var i = 0;i<arrIndex.length;i++)
		{
			document.getElementById('var' + arrIndex[i]).value = '';
		}
	}
	document.frmVars.isSubmit.value = "R";
	document.frmVars.submit();
}
function goGetValue(varIndex, varDefVars, varBase)
{
	Start(document.getElementById('var' + varIndex),'../SmallQuery.asp?source=olkrep&rsIndex=<%=Request("rsIndex")%>&varIndex=' + varIndex + '&s=' + varDefVars + varBase,550,250,'Yes','Yes');
}
</script>
<script language="javascript" src="../js_up_down.js"></script>
<body topmargin="0" leftmargin="0" onfocus="chkWin()" onbeforeunload="opener.clearWin();" <% If Request("isSubmit") = "" Then %>onload="reload('');"<% End If %>>
<form method="POST" action="viewRepVals.asp" name="frmVars" onsubmit="javascript:return doLoadDesc();" webbot-action="--WEBBOT-SELF--">
<table border="0" cellpadding="0" cellspacing="0" width="350" id="table3">
  	<tr>
		<td>
		<img src="images/spacer.gif" width="15" height="1" border="0" alt=""></td>
		<td>
		<img src="images/spacer.gif" width="318" height="1" border="0" alt=""></td>
		<td>
		<img src="images/spacer.gif" width="17" height="1" border="0" alt=""></td>
		<td><img src="images/spacer.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr>
		<td colspan="3">
		<img name="reportes_olkjpg_r1_c1" src="images/<%=Session("rtl")%>reportes_olkjpg_r1_c1.jpg" width="350" height="47" border="0" alt=""></td>
		<td>
		<img src="images/spacer.gif" width="1" height="47" border="0" alt=""></td>
	</tr>
	<tr>
		<td background="images/reportes_olkjpg_r2_c1.jpg">
		<img name="reportes_olkjpg_r2_c1" src="images/<%=Session("rtl")%>reportes_olkjpg_r2_c1.jpg" width="15" height="321" border="0" alt=""></td>
		<td   class="TblRep" valign="top">
		<table border="0" cellpadding="0" cellspacing="2" width="100%" id="table4">
			<tr>
				<td class="TblGreenTlt" colspan="2"><p align="center"><font face="verdana" color="black" size="1"><b><%=rs("rsName")%></b></font></p></td>
			</tr>
			<% If rsTop Then %>
			<tr class="TblGreenNrm">
				<td width="50%"><font face="verdana" color="black" size="1"><%=getviewRepValsLngStr("DtxtTop")%></font>&nbsp;</td>
				<td width="50%">
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td>
						<input id="varTop" name="varTop" value="<%=rsTopDef%>" size="7" onchange="javascript:chkNum(this, 'N');" onfocus="this.select();" maxlength="6"></td>
						<td valign="middle">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td><img src="../images/img_nud_up.gif" id="btnNewOrderUp"></td>
							</tr>
							<tr>
								<td><img src="images/spacer.gif"></td>
							</tr>
							<tr>
								<td><img src="../images/img_nud_down.gif" id="btnNewOrderDown"></td>
							</tr>
						</table>
						<script language="javascript">NumUDAttachMin('frmVars', 'varTop', 'btnNewOrderUp', 'btnNewOrderDown', 1);</script>
						</td>
					</tr>
				</table>
				</td>
			</tr>
			<% End If %>
			<%
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKGetRSVarsData" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@rsIndex") = Request("rsIndex")
			cmd("@LanID") = Session("LanID")
			rs.close
			rs.open cmd, , 3, 1
			set rQVal = Server.CreateObject("ADODB.RecordSet")
			set rBase = Server.CreateObject("ADODB.RecordSet")
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKGetRSVarsBase" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@LanID") = Session("LanID")
			cmd("@rsIndex") = Request("rsIndex")
			rBase.open cmd, , 3, 1
		Dim varSubmitDesc
		varSubmitDesc = ""
		do while not rs.eof
			enableControl = True
			If rs("varNotNull") = "Y" Then
				If notNullVars <> "" Then notNullVars = notNullVars & ", "
				notNullVars = notNullVars & "var" & rs("varIndex") & "~" & rs("varName") & "~" & rs("varType")
			End If
			If rs("varType") = "DD" or rs("varType") = "L" or rs("varType") = "CL" Then
				If varSubmitDesc <> "" Then varSubmitDesc = varSubmitDesc & ", "
				varSubmitDesc = varSubmitDesc & "var" & rs("varIndex") & "|" & rs("varType")
			End If
			Select Case rs("varType")
				Case "DD", "CL"
				   If rs("varDefVars") = "F" Then
				   		sql = "select T0.valValue, IsNull(T1.alterValText, T0.valText) valText " & _
				   				"from OLKRSVarsVals T0 " & _
				   				"left outer join OLKRSVarsValsAlterNames T1 on T1.rsIndex = T0.rsIndex and T1.varIndex = T0.varIndex and T1.valIndex = T0.valIndex and T1.LanID = " & Session("LanID") & " " & _
				   				"where T0.rsIndex = " & Request("rsIndex") & " and T0.varIndex = " & rs("varIndex")
				   Else
						sql = getSQL(true, rs("varQuery"))
				   End If
				Case "L"
				   If rs("varDefVars") = "F" Then
				   		sql = "select valValue, valText from OLKRSVarsVals where rsIndex = " & Request("rsIndex") & " and varIndex = " & rs("varIndex")
				   Else
				   		sql = getSQL(true, rs("varQuery"))
				   End If
				Case "Q"
					If rs("varDefVars") = "Q" Then
						sql = getSQL(false, rs("varQuery"))
					End If
			End Select
			If Request("isSubmit") <> "R" Then
				If rs("varType") <> "DP" and rs("DefValBy") = "V" Then
					defValue = rs("DefValValue")
				ElseIf rs("varType") = "DP" and rs("DefValBy") = "V" Then
					defValue = FormatDate(rs("DefValDate"), False)
				ElseIf rs("DefValBy") = "Q" Then
					sqlVal = getSQL(true, rs("DefValValue"))
					set rQVal = conn.execute(sqlVal)
					If not rQVal.eof then
						If rs("varType") = "DP" Then
							defValue = FormatDate(rQVal(0), False)
						Else
							defValue = CStr(rQVal(0))
						End If
					Else
						defValue = ""
					End If
					rQVal.close
				Else
					defValue = ""
				End If
			Else
				defValue = Request("var" & rs("varIndex"))
			End If %>
			<tr class="TblGreenNrm">
				<td width="50%">
				<table cellpadding="0" cellspacing="0" border="0" width="100%">
					<tr>
						<td><font face="verdana" color="black" size="1"><%=rs("varName")%></font></td>
						<% Select Case rs("varType") 
							Case "Q" %><td width="15"><img border="0" src="../images/<%=Session("rtl")%>flechaselec.gif" width="15" height="13" <% If enableControl Then %>onclick="javascript:goGetValue(<%=rs("varIndex")%>, '<%=rs("varDefVars")%>', '<%=sql%>');"<% End If %>></td>
						<% Case "DP" %>
							<td width="16"><img border="0" src="../images/cal.gif" width="16" height="16" id="btn<%=rs("varIndex")%>"></td>
						<% End Select %>
					</tr>
				</table>
				</td>
				<td width="50%"><% Select Case rs("varType")
			   Case "DD" 
			   If enableControl Then set rd = conn.execute(sql) %>
			   <select name="var<%=rs("varIndex")%>" size="1" <% If Not enableControl Then %>disabled<% End If %> <% If rs("IsBase") = "Y" Then %> onchange="reload(<%=rs("targetIndex")%>);"<% End If %> style="width: 96%;">
			   <% If enableControl Then %>
			   <option></option>
			   <% do while not rd.eof %>
			   <option <% If defValue = CStr(rd(0)) Then %>selected<% End If %> value="<%=rd(0)%>"><%=myHTMLEncode(rd(1))%></option>
			   <% rd.movenext
			   loop
			   Else %>
			   <option value=""><%=getviewRepValsLngStr("LtxtSelectEnter")%> "<%=selectName%>"</option>
			   <% End If %>
			   </select>
			<% Case "T" %><input type="text" name="var<%=rs("varIndex")%>" size="20" onchange="chkNum(this, '<%=rs("varDataType")%>');<% If rs("IsBase") = "Y" Then %>reload(<%=rs("targetIndex")%>);<% End If %>" value="<%=defValue%>" style="width: 96%; ">
			<% Case "L"
			   If enableControl Then set rd = conn.execute(sql) %>
				<select name="var<%=rs("varIndex")%>" size="5"  <% If Not enableControl Then %>disabled<% End If %> <% If rs("IsBase") = "Y" Then %> onchange="reload(<%=rs("targetIndex")%>);"<% End If %> style="width: 96%; "><%
			If enableControl Then
			   do while not rd.eof
			   %><option <% If defValue = CStr(rd(0)) Then %>selected<% End If %> value="<%=rd(0)%>"><%=myHTMLEncode(rd(1))%></option>
			   <% rd.movenext
			   loop
			   Else %>
			   <option value=""><%=getviewRepValsLngStr("LtxtSelectEnter")%> "<%=selectName%>"</option>
			   <% End If %>
				</select>
			<% Case "Q" 
			%><input readonly type="text" name="var<%=rs("varIndex")%>" id="var<%=rs("varIndex")%>" size="16" <% If enableControl Then %>onclick="javascript:goGetValue(<%=rs("varIndex")%>, '<%=rs("varDefVars")%>', '<%=sql%>');" value="<%=defValue%>" <% Else %> disabled value="<%=getviewRepValsLngStr("LtxtSelectEnter")%> &quot;<%=selectName%>&quot;" <% End If %> <% If rs("IsBase") = "Y" Then %> onchange="reload(<%=rs("targetIndex")%>);"<% End If %> style="width: 96%;">
			<% Case "DP" %><table border="0" cellpadding="0" cellspacing="0"><tr>
			<td>
			<input name="var<%=rs("varIndex")%>" id="var<%=rs("varIndex")%>" size="12" readonly  onclick="btn<%=rs("varIndex")%>.click()" value="<%=defValue%>" <% If rs("IsBase") = "Y" Then %> onchange="reload(<%=rs("targetIndex")%>);"<% End If %>></td></tr></table>
			<% Case "CL" 
			   If enableControl Then %>
				<ilayer name="scroll<%=rs("varIndex")%>1" width=100% height=120 clip="0,0,170,150">
				<layer name="scroll<%=rs("varIndex")%>2" width=100% height=120 bgColor="white">
				<div id="scroll<%=rs("varIndex")%>3" style="width:100%;height:120px;overflow:auto">
			   	<% 
			   	set rd = conn.execute(sql)
			   	i = 0
			   	do while not rd.eof %>
				<input type="checkbox" name="var<%=rs("varIndex")%>" value="'<%=Replace(myHTMLEncode(rd(0)), "'", "''")%>'" id="var<%=rs("varIndex")%><%=i%>" class="noborder" <% If rs("IsBase") = "Y" Then %> onclick="reload(<%=rs("targetIndex")%>);"<% End If %>><label id="txt<%=rs("varIndex")%><%=i%>" for="var<%=rs("varIndex")%><%=i%>"><%=myHTMLEncode(rd(1))%></label><br>
				<% i = i + 1
				rd.movenext
				loop
				Else %>
			   <label><%=getviewRepValsLngStr("LtxtSelectEnter")%></label> "<%=selectName%>"
				</div>
				</layer>
				</ilayer>
			   <% End If %>
			<% End Select %></td>
			</tr>
			<% rs.movenext
			loop %>
			<tr class="TblGreenNrm">
				<td width="50%" align="center">
				<input type="submit" name="btnQuery" value="<%=getviewRepValsLngStr("LtxtQuery")%>" >&nbsp;</td>
				<td width="50%" align="center">
				<input type="button" name="B2" value="<%=getviewRepValsLngStr("DtxtCancel")%>"  onclick="javascript:window.close()"></td>
			</tr>
		</table>
		</td>
		<td background="images/reportes_olkjpg_r2_c3.jpg">
		<img name="reportes_olkjpg_r2_c3" src="images/<%=Session("rtl")%>reportes_olkjpg_r2_c3.jpg" width="17" height="321" border="0" alt=""></td>
		<td>
		<img src="images/spacer.gif" width="1" height="321" border="0" alt=""></td>
	</tr>
	<tr>
		<td colspan="3">
		<img name="reportes_olkjpg_r3_c1" src="images/reportes_olkjpg_r3_c1.jpg" width="350" height="32" border="0" alt=""></td>
		<td>
		<img src="images/spacer.gif" width="1" height="32" border="0" alt=""></td>
	</tr>
</table>
<input type="hidden" name="pop" value="<%=Request("pop")%>">
<input type="hidden" name="AddPath" value="<%=Request("AddPath")%>">
<input type="hidden" name="rsIndex" value="<%=Request("rsIndex")%>">
<% 
If varSubmitDesc <> "" Then
	arrSubmitDesc = Split(varSubmitDesc, ", ")
	For i = 0 to UBound(arrSubmitDesc) %>
<input type="hidden" name="<%=Split(arrSubmitDesc(i), "|")(0)%>Desc" value="">
<%	Next
End If %>
<input type="hidden" name="isSubmit" value="Y">
</form>
<script type="text/javascript">
function doLoadDesc()
{
	<% If rsTop Then %>
	if (document.frmVars.varTop.value == '')
	{
		alert('<%=getviewRepValsLngStr("LtxtEnterFld")%>'.replace('{0}', '<%=getviewRepValsLngStr("DtxtTop")%>'));
		document.frmVars.varTop.focus();
		return false;
	}
	<% End If %>
	<% If notNullVars <> "" Then 
	ArrVal = Split(notNullVars, ", ")
	for i = 0 to UBound(ArrVal)
	ArrVal2 = Split(ArrVal(i),"~")
	If ArrVal2(2) <> "CL" Then  %>
	if (document.frmVars.<%=ArrVal2(0)%>.value == '')
	{
		alert('<%=getviewRepValsLngStr("LtxtEnterFld")%>'.replace('{0}', '<%=ArrVal2(1)%>'));
		if (!document.frmVars.<%=ArrVal2(0)%>.disabled) document.frmVars.<%=ArrVal2(0)%>.focus();
		return false;
	}
	<% Else %>
	if (!isChkboxChecked(document.frmVars.<%=ArrVal2(0)%>))
	{
		alert('<%=getviewRepValsLngStr("LtxtEnterFld")%>'.replace('{0}', '<%=ArrVal2(1)%>'));
		return false;
	}
	<% End If
	Next
	End If %>
	<% 
	If varSubmitDesc <> "" Then
	For i = 0 to UBound(arrSubmitDesc) 
	varID = Split(arrSubmitDesc(i), "|")(0)
	vType = Split(arrSubmitDesc(i), "|")(1)
	If vType <> "CL" Then
	%>
	document.frmVars.<%=varID%>Desc.value = document.frmVars.<%=varID%>.options[document.frmVars.<%=varID%>.selectedIndex].text;
	<% 
	Else %>
	document.frmVars.<%=varID%>Desc.value = getCheckListDesc(document.frmVars.<%=varID%>);
<%	End If 
	Next
	End If 
	%>
	return true;
}
function getCheckListDesc(Field)
{
	var retVal = '';
	if (Field)
	{
		if (Field.length)
		{
			for (var i = 0;i<Field.length;i++)
			{
				if (Field[i].checked)
				{
					if (retVal != '') retVal += ', ';
					retVal += document.getElementById(Field[i].id.replace('var', 'txt')).innerText;
				}
			}
		}
		else
		{
			if (Field.checked) retVal = document.getElementById(Field.id.replace('var', 'txt')).innerText;
		}
	}
	return retVal;
}
	<% 
	If rs.recordcount > 0 Then
	rs.movefirst 
	do while not rs.eof
		If rs("varType") = "DP" Then %>
	    Calendar.setup({
	        inputField     :    "var<%=rs("varIndex")%>",     // id of the input field
	        ifFormat       :    "<%=GetCalendarFormatString%>",      // format of the input field
	        button         :    "btn<%=rs("varIndex")%>",  // trigger for the calendar (button ID)
	        align          :    "Bl",           // alignment (defaults to "Bl")
	        singleClick    :    true
	    });
    <% 	End If
    rs.movenext
    loop
    End If %>
</script>

</body>
<% set rd = nothing
set rs = nothing
conn.close %>
<% Else  %>
<script language="javascript">
if (opener.document.frmGoRep.cmd) opener.document.frmGoRep.cmd.value = 'report';
opener.document.frmGoRep.target = '';
opener.document.frmGoRep.action = 'report.asp';
opener.document.frmGoRep.rsIndex.value='<%=Request("rsIndex")%>';
<% For each item in Request.Form
If item <> "rsIndex" and item <> "submit" Then %>
opener.document.frmGoRep.innerHTML += '<input type=\"hidden\" name=\"<%=item%>\" value=\"<%=Replace(Replace(Request(item), """", "\"""), "'", "\'")%>\">';
<% End If
Next %>
opener.document.frmGoRep.submit(); opener.clearWin(); window.close();</script>
<% End If %>
<%
Function getSQL(doQuery, Qry)
	If doQuery Then
		retVal= " declare @SlpCode int set @SlpCode = " & Session("vendid") & " "
	Else
		retVal = ""
	End If
	rBase.Filter = "varIndex = " & rs("varIndex")
	do while not rBase.eof
		If Request("var" & rBase("baseIndex")) <> "" Then
			If doQuery Then
				If rBase("varDataType") = "nvarchar" Then 
					MaxVar = "(" & rBase("varMaxChar") & ")"
				ElseIf rBase("varDataType") = "numeric" Then
					MaxVar = "(19,6)"
				Else
					MaxVar = ""
				End If
				retVal = retVal & "declare @" & rBase("varVar") & " " & rBase("varDataType") & " " & MaxVar & " "
				Select Case rBase("varDataType") 
					Case "nvarchar" 
						retVal = retVal & "set @" & rBase("varVar") & " = N'" & saveHTMLDecode(Request("var" & rBase("baseIndex")), False) & "' "
					Case "datetime"
						retVal = retVal & "set @" & rBase("varVar") & " = Convert(datetime,'" & SaveSqlDate(Request("var" & rBase("baseIndex"))) & "',120) "
					Case Else
						retVal = retVal & "set @" & rBase("VarVar") & " = " & Request("var" & rBase("baseIndex")) & " "
				End Select
			Else
				retVal = retVal & "&var" & rBase("baseIndex") & "=" & Request("var" & rBase("baseIndex"))
			End If
		Else
			selectName = rBase("varName")
			enableControl = False
			Exit Do
		End If
	rBase.movenext
	loop
	If doQuery Then retVal = retVal & " declare @LanID int set @LanID = " & Session("LanID") & " "
	If doQuery Then retVal = retVal & Qry
	retVal = QueryFunctions(retVal)
	getSQL = retVal
End Function %>
</html>
<!--#include file="lang/viewRepVals.asp" -->

<link rel="stylesheet" href="Reportes/style.css">
<% sql = "select IsNull(alterRSName, rsName) rsName, IsNull(alterRSDesc, rsDesc) rsDesc, rsTop, rsTopDef " & _
"from OLKRS T0 " & _
"left outer join OLKRSAlterNames T1 on T1.rsIndex = T0.rsIndex and T1.LanID = " & Session("LanID") & " " & _
"where T0.rsIndex = " & Request.Form("rsIndex")
set rs = conn.execute(sql)
set rd = Server.CreateObject("ADODB.RecordSet")
rsName = rs(0)
rsDesc = rs(1)
rsTop = rs("rsTop") = "Y"
rsTopDef = rs("rsTopDef")
rs.close %>
<script language="javascript">
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
function doCal(varId)
{
	document.frmViewRep.editVar.value = varId;
	document.frmViewRep.cmd.value = 'viewRepValsCal';
	document.frmViewRep.submit();
}
function doQuery(varId)
{
	document.frmViewRep.editVar.value = varId;
	document.frmViewRep.cmd.value = 'viewRepValsQry';
	document.frmViewRep.submit();
}
function doChkList(varId)
{
	document.frmViewRep.editVar.value = varId;
	document.frmViewRep.cmd.value = 'viewRepValsCL';
	document.frmViewRep.submit();
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
	//document.frmViewRep.action='viewRepVals.asp';
	document.frmViewRep.cmd.value = 'viewRepVals';
	document.frmViewRep.isSubmit.value = "R";
	document.frmViewRep.submit();
}

function MyIsNumeric(sText)
{
   if (sText == '') return false;
   
   var ValidChars = "0123456789.";
   var IsNumber=true;
   var Char;

 
   for (i = 0; i < sText.length && IsNumber == true; i++) 
      { 
      Char = sText.charAt(i); 
      if (ValidChars.indexOf(Char) == -1) 
         {
         IsNumber = false;
         }
      }
   return IsNumber;
   
}
</script>
<table border="0" cellspacing="0" width="100%" id="table1">
	<tr class="TblTltMnu">
		<td colspan="3"><img border="0" src="images/arrow_menu.gif" width="9" height="6">&nbsp;<%=rsName%>&nbsp;</td>
	</tr>
	<% If Not IsNull(rsDesc) Then %>
	<tr class="TblTltMnu">
		<td colspan="3"><%=rsDesc%>&nbsp;</td>
	</tr>
	<% End If %>
	<form method="POST" name="frmViewRep" action="operaciones.asp">
	<tr class="TblAfueraMnu">
		<td colspan="3">&nbsp;</td>
	</tr>
	<% If rsTop Then %>
	<tr class="TblAfueraMnu">
		<td><%=getviewRepValsLngStr("DtxtTop")%>&nbsp;</td>
		<td width="16"></td>
		<td>
		<input name="varTop" type="text" onchange="chkNum(this, 'int');" value="<%=rsTopDef%>" size="15">
		</td>
	</tr>
	<% End If %>
	<%
	sql = "select T0.varIndex, IsNull(T1.alterVarName, T0.varName) varName, T0.varVar, T0.varType, T0.varDataType, T0.varQuery, T0.varQueryField, T0.varMaxChar, T0.varNotNull, T0.varDefVars, T0.varShowRep, T0.DefValBy, T0.DefValValue, T0.DefValDate, " & _
		"Case When Exists(select 'A' from OLKRSVarsBase where rsIndex = T0.rsIndex and baseIndex = T0.varIndex) Then 'Y' Else 'N' End IsBase, " & _
		"OLKCommon.dbo.DBOLKGetRSVarTarget" & Session("ID") & "(T0.rsIndex, T0.varIndex) TargetIndex, DefValBy, DefValValue " & _
		"from OLKRSVars T0 " & _
		"left outer join OLKRSVarsAlterNames T1 on T1.rsIndex = T0.rsIndex and T1.varIndex = T0.varIndex and T1.LanID = " & Session("LanID") & " " & _
		"where T0.rsIndex = " & Request("rsIndex") & " order by T0.Ordr asc"
	rs.open sql, conn, 3, 1
	set rQVal = Server.CreateObject("ADODB.RecordSet")
	set rBase = Server.CreateObject("ADODB.RecordSet")
	sql = "select T1.varIndex, T1.baseIndex, varVar, IsNull(alterVarName, varName) varName, varDataType, varMaxChar " & _
			"from OLKRSVars T0 " & _
			"inner join OLKRSVarsBase T1 on T1.rsIndex = T0.rsIndex and T1.baseIndex = T0.varIndex " & _
			"left outer join OLKRSVarsAlterNames T2 on T2.rsIndex = T0.rsIndex and T2.varIndex = T0.varIndex and T2.LanID = " & Session("LanID") & " " & _
			"where T0.rsIndex = " & Request("rsIndex") & " "
	rBase.open sql, conn, 3, 1
	Dim varSubmitDesc
	varSubmitDesc = ""
	do while not rs.eof
		enableControl = True
		If rs("varNotNull") = "Y" Then
			If notNullVars <> "" Then notNullVars = notNullVars & ", "
			notNullVars = notNullVars & "var" & rs("varIndex") & "~" & rs("varName")
		End If
		If rs("varType") = "DD" or rs("varType") = "L" Then
			If varSubmitDesc <> "" Then varSubmitDesc = varSubmitDesc & ", "
			varSubmitDesc = varSubmitDesc & "var" & rs("varIndex")
		End If
		Select Case rs("varType")
			Case "DD", "L", "CL"
			   If rs("varDefVars") = "F" Then
			   		sql = "select valValue, valText " & _
			   				"from OLKRSVarsVals " & _
			   				"where rsIndex = " & Request("rsIndex") & " and varIndex = " & rs("varIndex")
			   Else
					sql = getSQL(true, rs("varQuery"))
			   End If
			'Case "L"
			 '  If rs("varDefVars") = "F" Then
			 '  		sql = "select valValue, valText from OLKRSVarsVals where rsIndex = " & Request("rsIndex") & " and varIndex = " & rs("varIndex")
			 '  Else
			 '  		sql = getSQL(true, rs("varQuery"))
			 '  End If
			Case "Q"
				If rs("varDefVars") = "Q" Then
					sql = getSQL(false, "")
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
			If rs("varType") = "CL" Then
				defValueDesc = Request("var" & rs("varIndex") & "Desc")
			End If
		End If %>
	<tr class="TblAfueraMnu">
		<td><%=Server.HTMLEncode(rs("varName"))%>&nbsp;</td>
		<td width="16"><% Select Case rs("varType") %>
			<% Case "Q" %><% If enableControl Then %><a href="javascript:doQuery(<%=rs("varIndex")%>);"><% End If %><img border="0" src="images/<%=Session("rtl")%>flechaselec.gif"><% If enableControl Then %></a><% End If %>
		<% Case "DP" %><a href="javascript:doCal(<%=rs("varIndex")%>);"><img border="0" src="images/cal.gif" width="16" height="16"></a>
		<% End Select %></td>
		<td>
		<% Select Case rs("varType") %>
		<% Case "DD"
		If enableControl Then set rd = conn.execute(sql) %>
		<select name="var<%=rs("varIndex")%>" size="1"  <% If Not enableControl Then %>disabled<% End If %> <% If rs("IsBase") = "Y" Then %> onchange="reload(<%=rs("targetIndex")%>);"<% End If %>>
		<% If enableControl Then %>
		<option></option>
		<% do while not rd.eof %><option <% If defValue = CStr(rd(0)) Then %>selected<% End If %> value="<%=rd(0)%>"><%=myHTMLEncode(rd(1))%></option>
		<% rd.movenext
		loop
		Else %>
		<option value=""><%=getviewRepValsLngStr("LtxtSelectEnter")%> "<%=selectName%>"</option>
		<% End If %>
		</select>
		<% Case "T" %><input name="var<%=rs("varIndex")%>" type="text" onchange="chkNum(this, '<%=rs("varDataType")%>');<% If rs("IsBase") = "Y" Then %>reload(<%=rs("targetIndex")%>);<% End If %>" value="<%=defValue%>" size="15">
		<% Case "L"
		If enableControl Then set rd = conn.execute(sql) %>
		<select name="var<%=rs("varIndex")%>" size="5" <% If Not enableControl Then %>disabled<% End If %> <% If rs("IsBase") = "Y" Then %> onchange="reload(<%=rs("targetIndex")%>);"<% End If %>><%
		If enableControl Then
		do while not rd.eof
		%><option <% If defValue = CStr(rd(0)) Then %>selected<% End If %> value="<%=rd(0)%>"><%=myHTMLEncode(rd(1))%></option>
		<% rd.movenext
		loop
		Else %>
		<option value=""><%=getviewRepValsLngStr("LtxtSelectEnter")%> "<%=selectName%>"</option>
		<% End If %>
		</select>
		<% Case "Q" %><input name="var<%=rs("varIndex")%>" type="text" readonly size="15" <% If enableControl Then %> value="<%=defValue%>" onclick="javascript:doQuery(<%=rs("varIndex")%>);"<% Else %> disabled value="Seleccione/Introduzca &quot;<%=selectName%>&quot;" <% End If %> <% If rs("IsBase") = "Y" Then %> onchange="reload(<%=rs("targetIndex")%>);"<% End If %>>
		<% Case "DP" %><input name="var<%=rs("varIndex")%>" type="text" readonly size="12" onclick="doCal(<%=rs("varIndex")%>);" value="<%=defValue%>" <% If rs("IsBase") = "Y" Then %> onchange="reload(<%=rs("targetIndex")%>);"<% End If %>>
		<% Case "CL" %><input name="var<%=rs("varIndex")%>Desc" type="text" readonly size="15" <% If enableControl Then %> value="<%=defValueDesc%>" onclick="javascript:doChkList(<%=rs("varIndex")%>);"<% Else %> disabled value="Seleccione/Introduzca &quot;<%=selectName%>&quot;" <% End If %> <% If rs("IsBase") = "Y" Then %> onchange="reload(<%=rs("targetIndex")%>);"<% End If %>>
		<input name="var<%=rs("varIndex")%>" type="hidden" <% If enableControl Then %> value="<%=defValue%>" <% End If %>>
		<% End Select %>
		</td>
	</tr>
	<% rs.movenext
	loop %>
	<tr class="TblAfueraMnu">
		<td colspan="3">&nbsp;</td>
	</tr>
	<tr class="TblAfueraMnu">
		<td colspan="3">
		<p align="center">
		<input type="submit" value="<%=getviewRepValsLngStr("DtxtAccept")%>" name="btnAccept" onclick="return valFrm();"></td>
	</tr>
		<input type="hidden" name="rsIndex" value="<%=Request("rsIndex")%>">
		<input type="hidden" name="editVar" value="">
		<input type="hidden" name="cmd" value="viewRep">
		<input type="hidden" name="isSubmit" value="Y">
	</form>
	<tr class="TblAfueraMnu">
		<td colspan="3">&nbsp;</td>
	</tr>
</table>
&nbsp;<script language="javascript">
function valFrm()
{
	<% If rsTop Then %>
	if (document.frmViewRep.varTop.value == '')
	{
		alert('<%=getviewRepValsLngStr("LtxtEnterFld")%>'.replace('{0}', '<%=getviewRepValsLngStr("DtxtTop")%>'));
		document.frmViewRep.varTop.focus();
		return false;
	}
	<% End If %>
	<% 
	If rs.RecordCount > 0 Then
	rs.movefirst
	do while not rs.eof
		If rs("varNotNull") = "Y" Then
			If rs.bookmark > 1 Then Response.write "else "
			If rs("varType") <> "DD" Then %>
			if (document.frmViewRep.var<%=rs("varIndex")%>.value == '')
			{
				alert('<%=getviewRepValsLngStr("LtxtEnterFld")%>'.replace('{0}', '<%=rs("varName")%>'));
				<% If rs("varType") <> "CL" Then %>if (!document.frmViewRep.var<%=rs("varIndex")%>.disabled) document.frmViewRep.var<%=rs("varIndex")%>.focus();<% End If %>
				return false;
			}
			<% Else %>
			if (document.frmViewRep.var<%=rs("varIndex")%>.value == '' || document.frmViewRep.var<%=rs("varIndex")%>.value == null)
			{
				alert('<%=getviewRepValsLngStr("LtxtSelFld")%>'.replace('{0}', '<%=rs("varName")%>'));
				if (!document.frmViewRep.var<%=rs("varIndex")%>.disabled) document.frmViewRep.var<%=rs("varIndex")%>.focus();
				return false;
			}
		<% 	End If
		End If
	rs.movenext
	loop
	End If %>
	document.frmViewRep.cmd.value='viewRep';
	return true;
}
</script>
<%
Function getSQL(doQuery, Qry)
	If doQuery Then
		retVal= "declare @LanID int set @LanID = " & Session("LanID") & " declare @SlpCode int set @SlpCode = " & Session("vendid") & " "
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
	If doQuery Then retVal = retVal & Qry
	getSQL = retVal
End Function %>
<% If Request("isSubmit") = "" Then %><script language="javascript">reload('');</script><% End If %>
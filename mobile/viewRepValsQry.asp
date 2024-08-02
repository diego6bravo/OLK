<!--#include file="lang/viewRepValsQry.asp" -->

<link rel="stylesheet" href="Reportes/style.css">
<% 
Select Case curCmd
	Case "viewRepValsQry"
		viewCmd = "viewRepVals"
		
		sql = "select IsNull(T2.alterRSName, T0.rsName) rsName, IsNull(T3.alterVarName, T1.varName) varName, T1.varDefVars, T1.varQuery, T1.varQueryField " & _
				"from OLKRS T0 " & _
				"inner join OLKRSVars T1 on T1.rsIndex = T0.rsIndex and T1.varIndex = " & Request("editVar") & " " & _
				"left outer join OLKRSAlterNames T2 on T2.rsIndex = T0.rsIndex and T2.LanID = " & Session("LanID") & " " & _
				"left outer join OLKRSVarsAlterNames T3 on T3.rsIndex = T0.rsIndex and T3.varIndex = T1.varIndex and T3.LanID = " & Session("LanID") & " " & _
				"where T0.rsIndex = " & Request.Form("rsIndex")
		set rs = conn.execute(sql)
		set rd = Server.CreateObject("ADODB.RecordSet")
		rsName = rs(0)
		varName = rs(1)
		varDefVars = rs(2)
		varQuery = rs(3)
		varQueryField = rs(4)
		rs.close
		
		If varDefVars = "Q" Then
			strQry = varQuery
		ElseIf varDefVars = "F" Then
			sql = "select 'select valValue As Value, valText As Text from OLKRSVarsVals where rsIndex = " & Request("rsIndex") & " and varIndex = " & Request("editVar") & "' SqlQuery from OLKRSVars where rsIndex = " & Request("rsIndex") & " and varIndex = " & Request("editVar")
			set rs = conn.execute(sql)
			strQry = rs(0)
			rs.close
		End If
		
		sqlSmall = ""
		set rBase = Server.CreateObject("ADODB.RecordSet")
		sql = "select varIndex, varVar, varName, varDataType, varMaxChar  " & _
				"from OLKRSVars T0  " & _
				"where T0.rsIndex = " & Request("rsIndex") & " and varIndex in " & _
				"(select baseIndex from OLKRSVarsBase where rsIndex = T0.rsIndex and varIndex = " & Request("editVar") & ") "
		set rBase = conn.execute(sql)
		do while not rBase.eof
			If rBase("varDataType") = "nvarchar" Then 
				MaxVar = "(" & rBase("varMaxChar") & ")"
			ElseIf rBase("varDataType") = "numeric" Then
				MaxVar = "(19,6)"
			Else
				MaxVar = ""
			End If
			sqlSmall = sqlSmall & "declare @" & rBase("varVar") & " " & rBase("varDataType") & " " & MaxChar & " "
			If rBase("varDataType") = "nvarchar" or rBase("varDataType") = "datetime" Then
				sqlSmall = sqlSmall & "set @" & rBase("varVar") & " = N'" & Request("var" & rBase("varIndex")) & "' "
			Else
				sqlSmall = sqlSmall & "set @" & rBase("VarVar") & " = " & Request("var" & rBase("varIndex")) & " "
			End If
		rBase.movenext
		loop
				
		sqlSmall = sqlSmall & "declare @LanID int set @LanID = " & Session("LanID") & " " & _
				"declare @SlpCode int set @SlpCode = " & Session("vendid") & " " & _
				"declare @branch int set @branch = " & Session("branch") & " " & _
				strQry
	Case "adSearchValsQry"
		viewCmd = "adSearch"
		sql = "select IsNull(T2.alterName, T0.Name) Name, IsNull(T3.alterName, T1.Name) VarName, T1.DefVars, T1.Query, T1.QueryField " & _
				"from OLKCustomSearch T0 " & _
				"inner join OLKCustomSearchVars T1 on T1.ObjectCode = T0.ObjectCode and T1.ID = T0.ID and T1.VarID = " & Request("editVar") & " " & _
				"left outer join OLKCustomSearchAlterNames T2 on T2.ObjectCode = T0.ObjectCode and T2.ID = T0.ID and T2.LanID = " & Session("LanID") & " " & _
				"left outer join OLKCustomSearchVarsAlterNames T3 on T3.ObjectCode = T0.ObjectCode and T3.ID = T0.ID and T3.VarID = T1.VarID and T3.LanID = " & Session("LanID") & " " & _
				"where T0.ObjectCode = " & Request("adObjID") & " and T0.ID = " & Request.Form("ID")
		set rs = conn.execute(sql)
		set rd = Server.CreateObject("ADODB.RecordSet")
		rsName = rs(0)
		varName = rs(1)
		varDefVars = rs(2)
		varQuery = rs(3)
		varQueryField = rs(4)
		rs.close
		
		If varDefVars = "Q" Then
			strQry = varQuery
		ElseIf varDefVars = "F" Then
			sql = "select 'select valValue As Value, valText As Text from OLKCustomSearchVarsVals where ObjectCode = " & Request("adObjID") & " and ID = " & Request("ID") & " and VarID = " & Request("editVar") & "' SqlQuery from OLKCustomSearchVars where ObjectCode = " & Request("adObjID") & " and ID = " & Request("ID") & " and VarID = " & Request("editVar")
			set rs = conn.execute(sql)
			strQry = rs(0)
			rs.close
		End If
		
		sqlSmall = ""
	
		set rBase = Server.CreateObject("ADODB.RecordSet")
		sql = "select VarID, Variable, Name, DataType, MaxChar  " & _
				"from OLKCustomSearchVars T0  " & _
				"where T0.ObjectCode = " & Request("adObjID") & " and T0.ID = " & Request("ID") & " and VarID in " & _
				"(select BaseID from OLKCustomSearchVarsBase where ObjectCode = T0.ObjectCode and ID = T0.ID and VarID = " & Request("editVar") & ") "
		set rBase = conn.execute(sql)
		do while not rBase.eof
			If rBase("DataType") = "nvarchar" Then 
				MaxVar = "(" & rBase("MaxChar") & ")"
			ElseIf rBase("DataType") = "numeric" Then
				MaxVar = "(19,6)"
			Else
				MaxVar = ""
			End If
			sqlSmall = sqlSmall & "declare @" & rBase("Variable") & " " & rBase("DataType") & " " & MaxChar & " "
			If rBase("DataType") = "nvarchar" or rBase("DataType") = "datetime" Then
				sqlSmall = sqlSmall & "set @" & rBase("Variable") & " = N'" & Request("var" & rBase("VarID")) & "' "
			Else
				sqlSmall = sqlSmall & "set @" & rBase("Variable") & " = " & Request("var" & rBase("VarID")) & " "
			End If
		rBase.movenext
		loop
				
		sqlSmall = sqlSmall & "declare @LanID int set @LanID = " & Session("LanID") & " " & _
				"declare @SlpCode int set @SlpCode = " & Session("vendid") & " " & _
				"declare @CardCode nvarchar(20) set @CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "' " & _
				"declare @branch int set @branch = " & Session("branch") & " " & _
				strQry
End Select

set rs = conn.execute(QueryFunctions(sqlSmall))
 %>

<script language="javascript">
function setVal(val)
{
	document.frmViewRep.var<%=Request("editVar")%>.value=val;
	document.frmViewRep.cmd.value='<%=viewCmd%>';
	document.frmViewRep.isSubmit.value = 'R';
	document.frmViewRep.submit();
}
</script>
<table border="0" cellspacing="0" width="100%" id="table1">
	<tr class="TblTltMnu">
		<td colspan="2"><img border="0" src="images/arrow_menu.gif" width="9" height="6">&nbsp;<%=rsName%> - <%=varName%></td>
	</tr>
	<form method="POST" name="frmViewRep" action="operaciones.asp">
	<tr class="TblAfueraMnu">
		<td colspan="2" align="center">
		<table border="0" cellpadding="0" width="100%" id="table2">
			<tr class="TblTltMnu">
			 <% For each Field in rs.Fields %>
				<td>
				<p align="center"><%=Server.HTMLEncode(Field.Name)%>&nbsp;</td>
			<% next %>
			</tr>
			<% do while not rs.eof
			fldValue = rs(CStr(varQueryField)) %>
			<tr class="TblAfueraMnu">
			  <% For each Field in rs.Fields  
			  varx = varx + 1 %>
				<td width="175">
				<p><% If Not IsNull(Field) Then %><a href="javascript:setVal('<%=Server.HTMLEncode(fldValue)%>')"><%=Server.HTMLEncode(Field)%></a><% End If %>&nbsp;</td>
			<% next %>
			</tr>
		  <% varx = 0
		  rs.movenext
		  loop %>
		</table>
		</td>
	</tr>
	<tr class="TblAfueraMnu">
		<td colspan="2">
		<p align="center">
		</td>
	</tr>
		<% 	For each itm in Request.Form %>
		<input type="hidden" name="<%=itm%>" value="<%=Request(itm)%>">
		<% 	Next %>
	<tr class="TblAfueraMnu">
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	</tr>
	<tr class="TblAfueraMnu">
		<td>
		<p align="left">&nbsp;</td>
		<td>
		<p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
		<input type="submit" name="btnCancel" value="<%=getviewRepValsQryLngStr("DtxtCancel")%>" onclick="javascript:document.frmViewRep.isSubmit.value = 'R';document.frmViewRep.cmd.value='<%=Request("cmd")%>';"></td>
	</tr>
	</form>
</table>
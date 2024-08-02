<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>

<!--#include file="chkLogin.asp" -->
<!--#include file="myHTMLEncode.asp"-->
<!--#include file="repVars.inc"-->
<% 
varx = Request.ServerVariables("URL")
saveDir = Server.MapPath(Mid(varx, 1, Len(varx)-22) & "temp/")
set cmd = Server.CreateObject("ADODB.Command")
FileName = Request("FileName")
set myXml = Server.CreateObject("MSXML2.DOMDocument")
myXml.async = False
myXml.Load(saveDir & "\" & FileName)

set myNodes = myXml.documentElement.selectNodes("/Reports/report")
For each nod in myNodes
	rsIndex = nod.selectSingleNode("rsIndex").text
	rsIndexVal = rsIndex
	rsIndex = Replace(rsIndex, "-", "_")
	If Request("rsIndex" & rsIndex) <> "" Then
	
		If Request("Active" & rsIndex) = "Y" Then Active = "Y" Else Active = "N"
		
		LoadCmd("OLKAdminRS")
		
		cmd("@rsName") 		= Request("rsName" & rsIndex)
		cmd("@rsQuery") 	= nod.selectSingleNode("rsQuery").text
		cmd("@rsTop")		= nod.selectSingleNode("rsTop").text
		cmd("@rsTopDef")	= nod.selectSingleNode("rsTopDef").text
		cmd("@rgIndex") 	= Request("rgIndex" & rsIndex)
		cmd("@Refresh") 	= nod.selectSingleNode("Refresh").text
		cmd("@Active") 		= Active
		
		cmd("@LinkOnly") 	= nod.selectSingleNode("LinkOnly").text
		cmd("@Action") 		= "A"
		
		If nod.selectSingleNode("rsDesc").text <> "" Then cmd("@rsDesc") = nod.selectSingleNode("rsDesc").text
		
		cmd.execute()
		
		newIndex = cmd("@rsIndex")
		
		set myAlterNames = nod.selectNodes("AlterNames/AlterName")
		For each alterName in myAlterNames
			sql = "insert OLKRSAlterNames(LanID, rsIndex, alterRSName, alterRSDesc) " & _
					"values(" & alterName.selectSingleNode("LanID").text & ", " & _
								newIndex & ", "
			If alterName.selectSingleNode("alterRSName").text <> "" Then
				sql = sql & "N'" & Replace(alterName.selectSingleNode("alterRSName").text , "'", "''")& "' "
			Else
				sql = sql & "NULL "
			End If
			
			sql = sql & ", "
			
			If alterName.selectSingleNode("alterRSDesc").text <> "" Then
				sql = sql & "N'" & Replace(alterName.selectSingleNode("alterRSDesc").text, "'", "''") & "'"
			Else
				sql = sql & "NULL"
			End If
			
			sql = sql & ")"
			
			conn.execute(sql)
		Next
		
		set myVars = nod.selectNodes("Variables/Variable")
		For each var in myVars
			varIndex = var.selectSingleNode("varIndex").text
			
			LoadCmd("OLKAdminRSVars")
			cmd("@rsIndex") 		= newIndex
			cmd("@varIndex") 		= varIndex
			cmd("@varName")			= var.selectSingleNode("varName").text
			cmd("@varVar") 			= var.selectSingleNode("varVar").text
			cmd("@varType") 		= var.selectSingleNode("varType").text
			cmd("@varDataType") 	= var.selectSingleNode("varDataType").text
			cmd("@varNotNull") 		= var.selectSingleNode("varNotNull").text
			cmd("@varShowRep") 		= var.selectSingleNode("varShowRep").text
			cmd("@DefValBy") 		= var.selectSingleNode("DefValBy").text
			cmd("@Ordr")			= var.selectSingleNode("Ordr").text
			cmd("@Action") 			= "A"
			If var.selectSingleNode("varQuery").text		<> "" Then 		cmd("@varQuery") 		= var.selectSingleNode("varQuery").text
			If var.selectSingleNode("varQueryField").text 	<> "" Then 		cmd("@varQueryField") 	= var.selectSingleNode("varQueryField").text
			If var.selectSingleNode("varMaxChar").text 		<> "" Then 		cmd("@varMaxChar") 		= var.selectSingleNode("varMaxChar").text
			If var.selectSingleNode("varDefVars").text 		<> "" Then 		cmd("@varDefVars") 		= var.selectSingleNode("varDefVars").text
			If var.selectSingleNode("DefValValue").text 	<> "" Then 		cmd("@DefValValue") 	= var.selectSingleNode("DefValValue").text
			If var.selectSingleNode("DefValDate").text 		<> "" Then 		cmd("@DefValDate") 		= var.selectSingleNode("DefValDate").text
			cmd.execute()
			
			set myAlterNames = var.selectNodes("AlterNames/AlterName")
			For each alterName in myAlterNames
				sql = "insert OLKRSVarsAlterNames(LanID, rsIndex, varIndex, alterVarName) " & _
						"values(" & alterName.selectSingleNode("LanID").text & ", " & _
									newIndex & ", " & _
									alterName.selectSingleNode("varIndex").text & ", "
									
				If alterName.selectSingleNode("alterVarName").text <> "" Then
					sql = sql & "N'" & Replace(alterName.selectSingleNode("alterVarName").text, "'", "''") & "'"
				Else
					sql = sql & "NULL"
				End If
				
				sql = sql & ")"
				
				conn.execute(sql)
			Next
			
			set myBases = var.selectNodes("VariableBases/VariableBase")
			For each base in myBases
				LoadCmd("OLKAdminRSVarsBase")
				cmd("@rsIndex") 	= newIndex
				cmd("@varIndex") 	= varIndex
				cmd("@baseIndex")	= base.selectSingleNode("baseIndex").text
				cmd.execute()
			Next
			
			set myVarsVals = var.selectNodes("VariableValues/VariableValue")
			For each val in myVarsVals
				LoadCmd("OLKAdminRSVarsVals")
				cmd("@rsIndex") 	= newIndex
				cmd("@varIndex") 	= varIndex
				cmd("@valValue")	= val.selectSingleNode("valValue").text
				cmd("@valText")		= val.selectSingleNode("valText").text
				cmd.execute()
				
				set myAlterNames = val.selectNodes("AlterNames/AlterName")
				For each alterName in myAlterNames
					sql = "insert OLKRSVarsValsAlterNames(LanID, rsIndex, varIndex, valIndex, alterValText) " & _
							"values(" & alterName.selectSingleNode("LanID").text & ", " & _
										newIndex & ", " & _
										alterName.selectSingleNode("varIndex").text & ", " & _
										alterName.selectSingleNode("valIndex").text & ", "
										
					If alterName.selectSingleNode("alterValText").text <> "" Then
						sql = sql & "N'" & Replace(alterName.selectSingleNode("alterValText").text, "'", "''") & "'"
					Else
						sql = sql & "NULL"
					End If
					
					sql = sql & ")"
					
					conn.execute(sql)
				Next
			Next
			
		Next
		
		set myTotals = nod.selectNodes("Totals/Total")
		For each total in myTotals
			LoadCmd("OLKAdminRSTotals")
			cmd("@rsIndex") 	= newIndex
			cmd("@colName") 	= total.selectSingleNode("colName").text
			cmd("@colTotal")	= total.selectSingleNode("colTotal").text
			cmd("@colAlign")	= total.selectSingleNode("colAlign").text
			cmd("@colFormat")	= total.selectSingleNode("colFormat").text
			cmd("@colSum")		= total.selectSingleNode("colSum").text
			cmd("@colNB")		= total.selectSingleNode("colNB").text
			cmd("@colShow")		= total.selectSingleNode("colShow").text
			cmd("@linkType")	= total.selectSingleNode("linkType").text
			cmd("@linkPopup")	= total.selectSingleNode("linkPopup").text
			cmd("@linkCat")		= total.selectSingleNode("linkCat").text
			cmd("@Action") 		= "A"
			If total.selectSingleNode("linkObject").text 	<> "" Then cmd("@linkObject")	= total.selectSingleNode("linkObject").text
			If total.selectSingleNode("linkLink").text 		<> "" Then cmd("@linkLink")		= total.selectSingleNode("linkLink").text
			cmd.execute()
			
			set myAlterNames = total.selectNodes("AlterNames/AlterName")
			For each alterName in myAlterNames
				sql = "insert OLKRSTotalsAlterNames(LanID, rsIndex, colName, alterColName) " & _
						"values(" & alterName.selectSingleNode("LanID").text & ", " & _
									newIndex & ", " & _
									"N'" & Replace(alterName.selectSingleNode("colName").text, "'", "''") & "', "
									
				If alterName.selectSingleNode("alterColName").text <> "" Then
					sql = sql & "N'" & Replace(alterName.selectSingleNode("alterColName").text, "'", "''") & "'"
				Else
					sql = sql & "NULL"
				End If
				
				sql = sql & ")"
				
				conn.execute(sql)
			Next
		Next
		
		set myLinksVars = nod.selectNodes("LinksVars/LinksVar")
		For each var in myLinksVars
			LoadCmd("OLKAdminRSLinksVars")
			cmd("@rsIndex")		= newIndex
			cmd("@colName") 	= var.selectSingleNode("colName").text
			cmd("@varId")		= var.selectSingleNode("varId").text
			cmd("@valBy")		= var.selectSingleNode("valBy").text
			If var.selectSingleNode("valValue").text 	<> "" Then cmd("@valValue") 	= var.selectSingleNode("valValue").text
			If var.selectSingleNode("valValDat").text 	<> "" Then cmd("@valValDat") 	= var.selectSingleNode("valValDat").text
			cmd.execute()
		next
		
		set myColors = nod.selectNodes("Colors/Color")
		For each color in myColors
			LoadCmd("OLKAdminRSColors")
			cmd("@rsIndex")			= newIndex
			cmd("@ColorID")			= color.selectSingleNode("ColorID").text
			cmd("@LineID")			= color.selectSingleNode("LineID").text
			cmd("@Alias")			= color.selectSingleNode("Alias").text
			cmd("@colName")			= color.selectSingleNode("colName").text
			cmd("@colType")			= color.selectSingleNode("colType").text
			cmd("@colOp")			= color.selectSingleNode("colOp").text
			cmd("@colOpBy")			= color.selectSingleNode("colOpBy").text
			cmd("@FontBold")		= color.selectSingleNode("FontBold").text
			cmd("@FontItalic")		= color.selectSingleNode("FontItalic").text
			cmd("@FontUnderline")	= color.selectSingleNode("FontUnderline").text
			cmd("@FontStrike")		= color.selectSingleNode("FontStrike").text
			cmd("@FontBlink")		= color.selectSingleNode("FontBlink").text
			cmd("@ApplyToRow")		= color.selectSingleNode("ApplyToRow").text
			cmd("@Active")			= color.selectSingleNode("Active").text
			cmd("@Ordr")			= color.selectSingleNode("Ordr").text
			cmd("@Ordr2")			= color.selectSingleNode("Ordr2").text
			
			If color.selectSingleNode("colValue").text 		<> "" Then cmd("@colValue") 	= color.selectSingleNode("colValue").text
			If color.selectSingleNode("colValDate").text 	<> "" Then cmd("@colValDate") 	= color.selectSingleNode("colValDate").text
			If color.selectSingleNode("FontFace").text 		<> "" Then cmd("@FontFace") 	= color.selectSingleNode("FontFace").text
			If color.selectSingleNode("FontSize").text 		<> "" Then cmd("@FontSize") 	= color.selectSingleNode("FontSize").text
			If color.selectSingleNode("ForeColor").text 	<> "" Then cmd("@ForeColor") 	= color.selectSingleNode("ForeColor").text
			If color.selectSingleNode("BackColor").text 	<> "" Then cmd("@BackColor") 	= color.selectSingleNode("BackColor").text
			If color.selectSingleNode("ApplyToCol").text 	<> "" Then cmd("@ApplyToCol") 	= color.selectSingleNode("ApplyToCol").text
			cmd("@Action") = "A"
			cmd.execute()
		Next
		
		ClearTableData "RSNewRepLnkClear", newIndex, ""
	End If
Next

%>
<html>
<head></head>
<body>
<script language="javascript">
opener.location.href='adminReps.asp?uType=<%=Request("UserType")%>';
window.close();
</script>
</body>
</html>
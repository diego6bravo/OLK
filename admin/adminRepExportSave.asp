<%
Response.Buffer = True
Dim strFilePath, strFileSize, strFileName
%>
<!--#include file="chkLogin.asp" -->
<!--#include file="myHTMLEncode.asp"-->
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<% 
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="repVars.inc"--><% 
strFileName = repTbl & "Reps.xml"
Response.AddHeader "Content-Disposition", "attachment; filename=" & strFileName
Response.Charset = "UTF-8"
Response.ContentType = "file/Folder"

sql = "select * from " & repTbl & "RS where rsIndex in (" & Request("rsIndex") & ")"
set rs = Server.CreateObject("ADODB.RecordSet")
rs.open sql, conn, 3, 1

sql = "select * from " & repTbl & "RSAlterNames where rsIndex in (" & Request("rsIndex") & ")"
set rsAlterNames = Server.CreateObject("ADODB.RecordSet")
rsAlterNames.open sql, conn, 3, 1

sql = "select * from " & repTbl & "RSVars where rsIndex in (" & Request("rsIndex") & ")"
set rsVars = Server.CreateObject("ADODB.RecordSet")
rsVars.open sql, conn, 3, 1

sql = "select * from " & repTbl & "RSVarsAlterNames where rsIndex in (" & Request("rsIndex") & ")"
set rsVarsAlterNames = Server.CreateObject("ADODB.RecordSet")
rsVarsAlterNames.open sql, conn, 3, 1

sql = "select * from " & repTbl & "RSVarsBase where rsIndex in (" & Request("rsIndex") & ")"
set rsVarsBase = Server.CreateObject("ADODB.RecordSet")
rsVarsBase.open sql, conn, 3, 1

sql = "select * from " & repTbl & "RSVarsVals where rsIndex in (" & Request("rsIndex") & ")"
set rsVarsVals = Server.CreateObject("ADODB.RecordSet")
rsVarsVals.open sql, conn, 3, 1

sql = "select * from " & repTbl & "RSVarsValsAlterNames where rsIndex in (" & Request("rsIndex") & ")"
set rsVarsValsAlterNames = Server.CreateObject("ADODB.RecordSet")
rsVarsValsAlterNames.open sql, conn, 3, 1

sql = "select * from " & repTbl & "RSTotals where rsIndex in (" & Request("rsIndex") & ")"
set rsTotals = Server.CreateObject("ADODB.RecordSet")
rsTotals.open sql, conn, 3, 1

sql = "select * from " & repTbl & "RSTotalsAlterNames where rsIndex in (" & Request("rsIndex") & ")"
set rsTotalsAlterNames = Server.CreateObject("ADODB.RecordSet")
rsTotalsAlterNames.open sql, conn, 3, 1

sql = "select * from " & repTbl & "RSLinksVars where rsIndex in (" & Request("rsIndex") & ")"
set rsLinksVars = Server.CreateObject("ADODB.RecordSet")
rsLinksVars.open sql, conn, 3, 1

sql = "select * from " & repTbl & "RSColors where rsIndex in (" & Request("rsIndex") & ")"
set rsColors = Server.CreateObject("ADODB.RecordSet")
rsColors.open sql, conn, 3, 1

Response.Write "<?xml version=""1.0""?>" & VbNewLine
Response.Write "<Reports>" & VbNewLine

do while not rs.eof
	Response.Write "	<report id=""" & rs("rsIndex") & """  name=""" & doEncode(rs("rsName")) & """>" & VbNewLine
	For each fld in rs.Fields
		Response.Write "		<" & doEncode(fld.Name) & ">" & doEncode(fld) & "</" & doEncode(fld.Name) & ">" & VbNewLine
	Next
	
	rsAlterNames.Filter = "rsIndex = " & rs("rsIndex")
	Response.Write "		<AlterNames>" & VbNewLine
		do while not rsAlterNames.eof
			Response.Write "			<AlterName>" & VbNewLine
				For each fld in rsAlterNames.Fields
				Response.Write "			<" & doEncode(fld.Name) & ">" & doEncode(fld) & "</" & doEncode(fld.Name) & ">" & VbNewLine
				Next
			Response.Write "			</AlterName>" & VbNewLine
		rsAlterNames.movenext
		loop
	Response.Write "		</AlterNames>" & VbNewLine
	
	rsVars.Filter = "rsIndex = " & rs("rsIndex")
	Response.Write "		<Variables>" & VbNewLine
	If Not rsVars.Eof Then
		do while not rsVars.eof
			Response.Write "			<Variable id=""" & rsVars("varIndex") & """>" & VbNewLine
			For each fld in rsVars.Fields
				Response.Write "			<" & doEncode(fld.Name) & ">" & doEncode(fld) & "</" & doEncode(fld.Name) & ">" & VbNewLine
			Next
			
			rsVarsAlterNames.Filter = "rsIndex = " & rs("rsIndex") & " and varIndex = " & rsVars("varIndex")
			Response.Write "			<AlterNames>" & VbNewLine
				do while not rsVarsAlterNames.eof
					Response.Write "			<AlterName>" & VbNewLine
					For each fld in rsVarsAlterNames.Fields
						Response.Write "			<" & doEncode(fld.Name) & ">" & doEncode(fld) & "</" & doEncode(fld.Name) & ">" & VbNewLine
					Next
					Response.Write "			</AlterName>" & VbNewLine
				rsVarsAlterNames.movenext
				loop
			Response.Write "			</AlterNames>" & VbNewLine
						
			
			rsVarsBase.Filter = "rsIndex = " & rs("rsIndex") & " and varIndex = " & rsVars("varIndex")
			Response.Write "			<VariableBases>" & VbNewLine
			If Not rsVarsBase.eof then
				do while not rsVarsBase.eof
					Response.Write "			<VariableBase id=""" & rsVarsBase("baseIndex") & """>" & VbNewLine
					For each fld in rsVarsBase.Fields
						Response.Write "				<" & doEncode(fld.Name) & ">" & doEncode(fld) & "</" & doEncode(fld.Name) & ">" & VbNewLine
					Next
					Response.Write "			</VariableBase>" & VbNewLine
				rsVarsBase.movenext
				loop
			End If
			Response.Write "			</VariableBases>" & VbNewLine
			
			rsVarsVals.Filter = "rsIndex = " & rs("rsIndex") & " and varIndex = " & rsVars("varIndex")
			Response.Write "			<VariableValues>" & VbNewLine
			If Not rsVarsVals.Eof Then
				do while not rsVarsVals.eof
					Response.Write "			<VariableValue id=""" & rsVarsVals("valIndex") & """>" & VbNewLine
					For each fld in rsVarsVals.Fields
						Response.Write "				<" & doEncode(fld.Name) & ">" & doEncode(fld) & "</" & doEncode(fld.Name) & ">" & VbNewLine
					Next
					rsVarsValsAlterNames.Filter = "rsIndex = " & rs("rsIndex") & " and varIndex = " & rsVarsVals("varIndex") & " and valIndex = " & rsVarsVals("valIndex")
					Response.Write "			<AlterNames>" & VbNewLine
						do while not rsVarsValsAlterNames.eof
							Response.Write "			<AlterName>" & VbNewLine
							For each fld in rsVarsValsAlterNames.Fields
								Response.Write "			<" & doEncode(fld.Name) & ">" & doEncode(fld) & "</" & doEncode(fld.Name) & ">" & VbNewLine
							Next
							Response.Write "			</AlterName>" & VbNewLine
						rsVarsValsAlterNames.movenext
						loop
					Response.Write "			</AlterNames>" & VbNewLine
					Response.Write "			</VariableValue>" & VbNewLine
				rsVarsVals.movenext
				loop
			End If
			Response.Write "			</VariableValues>" & VbNewLine
			
			Response.Write "			</Variable>" & VbNewLine
		rsVars.movenext
		loop
		
	End If
	Response.Write "		</Variables>" & VbNewLine
	
	rsTotals.Filter = "rsIndex = " & rs("rsIndex")
	Response.Write "		<Totals>" & VbNewLine
	If Not rsTotals.eof Then
		do while not rsTotals.eof
			Response.Write "			<Total columnID=""" & rsTotals("colName") & """>" & VbNewLine
			For each fld in rsTotals.Fields
				Response.Write "				<" & doEncode(fld.Name) & ">" & doEncode(fld) & "</" & doEncode(fld.Name) & ">" & VbNewLine
			next
			
			rsTotalsAlterNames.Filter = "rsIndex = " & rs("rsIndex") & " and colName = '" & Replace(rsTotals("colName"), "'", "''") & "'"
			Response.Write "				<AlterNames>" & VbNewLine
			do while not rsTotalsAlterNames.eof
				Response.Write "					<AlterName>" & VbNewLine
				For each fld in rsTotalsAlterNames.Fields
					Response.Write "						<" & doEncode(fld.Name) & ">" & doEncode(fld) & "</" & doEncode(fld.Name) & ">" & VbNewLine
				Next
				Response.Write "					</AlterName>" & VbNewLine
			rsTotalsAlterNames.movenext
			loop
			Response.Write "				</AlterNames>" & VbNewLine
	
			Response.Write "			</Total>" & VbNewLine
		rsTotals.movenext
		loop
	End If
	Response.Write "		</Totals>" & VbNewLine
	
	rsLinksVars.Filter = "rsIndex = " & rs("rsIndex")
	Response.Write "		<LinksVars>" & VbNewLine
	If Not rsLinksVars.eof Then
		do while not rsLinksVars.eof
			Response.Write "			<LinksVar id=""" & rsLinksVars("varId") & """>" & VbNewLine
				For each fld in rsLinksVars.Fields
					Response.Write "				<" & doEncode(fld.Name) & ">" & doEncode(fld) & "</" & doEncode(fld.Name) & ">" & VbNewLine
				Next
			Response.Write "			</LinksVar>" & VbNewLine
		rsLinksVars.movenext
		loop
	End If
	Response.Write "		</LinksVars>" & VbNewLine
	
	rsColors.Filter = "rsIndex = " & rs("rsIndex")
	Response.Write "		<Colors>" & VbNewLine
	If not rsColors.Eof Then
		do while not rsColors.eof
			Response.Write "			<Color colorID=""" & rsColors("ColorID") & """ lineID=""" & rsColors("LineID") & """>" & VbNewLine
			For each fld in rsColors.Fields
					Response.Write "				<" & doEncode(fld.Name) & ">" & doEncode(fld) & "</" & doEncode(fld.Name) & ">" & VbNewLine
			next
			Response.Write "			</Color>" & VbNewLine
		rsColors.movenext
		loop
	End If
	Response.Write "		</Colors>" & VbNewLine

	Response.Write "	</report>" & VbNewLine
rs.movenext
loop

Response.Write "</Reports>"

set rs = nothing
conn.close

Response.Flush

Function doEncode(fld)
	retVal = fld
	If Not IsNull(fld) Then retVal = Server.HTMLEncode(fld)
	doEncode = retVal
End Function
%>
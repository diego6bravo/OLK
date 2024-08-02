<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>

<!--#include file="Upload/ShadowUploader.asp" -->
<!--#include file="chkLogin.asp" -->
<!--#include file="myHTMLEncode.asp"-->
<!--#include file="lang/adminRepImport.asp" -->
<!--#include file="repVars.inc"--><% 

Dim rsVars
set rsVars = Server.CreateObject("ADODB.RecordSet")
Dim ErrMsg
If Request("doUpload") = "Y" Then
	set rd = Server.CreateObject("ADODB.RecordSet")
	sql = 	"select rgIndex, rgName " & _
			"from " & repTbl & "RG where UserType = '" & Request("UserType") & "' and rgIndex >= 0 " & _
			"order by rgName asc"
	rd.open sql, conn, 3, 1
	
End If

varx = Request.ServerVariables("URL")
saveDir = Server.MapPath(Mid(varx, 1, Len(varx)-18) & "temp/")

%>
<html <% if session("rtl") <> "" then %>dir="rtl" <% end if %>>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=getadminRepImportLngStr("LttlRepImp")%></title>

<link rel="stylesheet" type="text/css" title="admin" href="style/style_admin_<%=Session("style")%>.css">
<script type="text/javascript">
var ignoreUnload = false;
var txtValXmlFile = '<%=getadminRepImportLngStr("LtxtValXmlFile")%>';
var txtValSelRep = '<%=getadminRepImportLngStr("LtxtValSelRep")%>';
var txtValRepName = '<%=getadminRepImportLngStr("LtxtValRepName")%>';
var txtValRepGrp = '<%=getadminRepImportLngStr("LtxtValRepGrp")%>';
</script>
<script type="text/javascript" src="adminRepImport.js"></script>
<script type="text/javascript" src="general.js"></script>
</head>

<body bgcolor="#F9FDFF" marginwidth="0" marginheight="0" topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" onbeforeunload="if(!ignoreUnload)opener.clearWin();">

<% If Request("doUpload") <> "Y" Then %>
<form method="POST" enctype="multipart/form-data" action="adminRepImport.asp?doUpload=Y&UserType=<%=Request("UserType")%>&pop=Y" onsubmit="return valFrmUpload(this);">
<table border="0" cellpadding="0" width="100%" id="table1" style="font-family: Verdana; font-size: 10px">
		<tr>
			<td class="TblRepTlt" colspan="2"><%=getadminRepImportLngStr("LttlRepImp")%></td>
		</tr>
		<tr class="TblRepTlt">
			<td>
			<%=getadminRepImportLngStr("LtxtSelDataFile")%> (XML)</td>
			<td>
			<p align="right">
			<input type="file" name="xmlFile" size="58" onchange="if (this.value.substring(this.value.length-3).toLowerCase() != 'xml'){ alert('<%=getadminRepImportLngStr("LtxtValXmlFile")%>');this.value=''; };"></td>
		</tr>
		<tr>
			<td>
			&nbsp;</td>
			<td>
			<p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
			<input type="submit" value="<%=getadminRepImportLngStr("LtxtNext")%>" name="btnNext"></td>
		</tr>
</table>
</form>
		<% Else %>
<form method="POST" name="frmImport" onsubmit="return valFrmImp()" action="adminRepImportSave.asp">
<table border="0" cellpadding="0" width="100%" id="table1" style="font-family: Verdana; font-size: 10px">
		<tr>
			<td class="TblRepTlt" colspan="2"><%=getadminRepImportLngStr("LttlRepImp")%></td>
		</tr><% 
		Dim objUpload
		set objUpload = New ShadowUpload
		If objUpload.GetError <> "" Then %>
		<tr>
			<td colspan="2">
			<%="" & getadminRepImportLngStr("LtxtErrUpdFile") & ": " & objUpload.GetError %>
			</td>
		</tr>
		<% Else %>
		<tr class="TblRepTbl">
			<td colspan="2">
			<img src="images/lentes.gif">&nbsp;<%=getadminRepImportLngStr("LtxtSelReps")%></td>
		</tr>
		<tr>
			<td colspan="2">
			<table border="0" cellpadding="0" cellspacing="1" width="100%" id="table2">
				<tr class="TblRepTlt">
					<td width="10">&nbsp;</td>
					<td align="center"><%=getadminRepImportLngStr("DtxtName")%></td>
					<td align="center"><%=getadminRepImportLngStr("DtxtGroup")%></td>
					<td align="center"><%=getadminRepImportLngStr("DtxtActive")%></td>
					<td align="center"><%=getadminRepImportLngStr("DtxtValid")%></td>
				</tr>
				<% 
				allDisabled = "disabled"
				FileName = objUpload.File(0).FileName
				Call objUpload.File(0).SaveToDisk(saveDir, "")
				
				set myXml = Server.CreateObject("MSXML2.DOMDocument")
				myXml.async = False
				myXml.Load(saveDir & "\" & FileName)
				
				set myNodes = myXml.documentElement.selectNodes("/Reports/report")
				For each nod in myNodes
				IsValid = True
				rsIndex = nod.selectSingleNode("rsIndex").text
				
				set myVars = nod.selectNodes("Variables/Variable")
				loadRSVars()
				sql = ""
				If Request.QueryString("UserType") = "C" Then sql = "declare @CardCode nvarchar(15) set @CardCode = '' "
				sql = sql & "declare @LanID int set @LanID = " & Session("LanID") & " "
				For each var in myVars
					sql = sql & " declare @" & var.selectSingleNode("varVar").text & " " & var.selectSingleNode("varDataType").text
					Select Case var.selectSingleNode("varDataType").text 
						Case "nvarchar"
							sql = sql & "(" & var.selectSingleNode("varMaxChar").text & ") "
						Case "numeric"
							sql = sql & "(19,6) "
					End Select
					
					If IsValid Then
						If var.selectSingleNode("varDefVars").text = "Q" Then
							set myBases = var.selectNodes("VariableBases/VariableBase")
							baseIndex = ""
							For each base in myBases
								If baseIndex <> "" Then baseIndex = baseIndex & " or "
								baseIndex = baseIndex & "varIndex = " & base.selectSingleNode("baseIndex").text
							Next
							rsVars.Filter = baseIndex
							varQuery = ""
							If Request.QueryString("UserType") = "C" Then varQuery = "declare @CardCode nvarchar(15) set @CardCode = '' "
							varQuery = varQuery & "declare @LanID int set @LanID = " & Session("LanID") & " "
							do while not rsVars.eof
								varQuery = varQuery & " declare @" & rsVars("varVar") & " " & rsVars("varDataType")
								Select Case rsVars("varDataType")
									Case "nvarchar"
										varQuery = varQuery & "(" & rsVars("varMaxChar") & ") "
									Case "numeric"
										varQuery = varQuery & "(19,6) "
								End Select
							rsVars.movenext
							loop
							varQuery = varQuery & " " & var.selectSingleNode("varQuery").text
							IsValid = ValidateQuery(varQuery, "V")
						End If
					
					
						If var.selectSingleNode("DefValBy").text = "Q" Then
							If IsValid Then IsValid = ValidateQuery(var.selectSingleNode("DefValValue").text, "D")
						End If 
					End If 
				Next
				sql = sql & " " & nod.selectSingleNode("rsQuery").text
				If nod.selectSingleNode("rsTop").text = "Y" Then sql = Replace(sql, "@top", 1)
				If IsValid Then IsValid = ValidateQuery(QueryFunctions(sql), "Q") 
				rsIndexVal = rsIndex
				rsIndex = Replace(rsIndex, "-", "_")
				%>
				<tr class="TblRepTbl">
					<td width="10">
					<p align="left">
					<input type="checkbox" name="rsIndex<%=rsIndex%>" id="rsIndex" value="<%=rsIndexVal%>" class="noborder" onclick="javascript:chkCheckAll('rsIndex');"></td>
					<td>
					<input type="text" name="rsName<%=rsIndex%>" value="<%=nod.selectSingleNode("rsName").text%>" size="50">&nbsp;<img src="images/undo.gif" border="0" alt="<%=getadminRepImportLngStr("LtxtRestoreName")%>" onclick="javascript:document.frmImport.rsName<%=rsIndex%>.value='<%=Replace(nod.selectSingleNode("rsName").text, "'", "\'")%>'"></td>
					<td>
					<select name="rgIndex<%=rsIndex%>" size="1">
					<option value=""></option>
					<% If rd.recordcount > 0 Then rd.MoveFirst
					do while not rd.EOF %>
					<option value="<%=rd("rgIndex")%>"><%=rd("rgName")%></option>
					<% rd.movenext
					loop %>
					</select>
					</td>
					<td>
					<p align="left">
					<input type="checkbox" id="Active" name="Active<%=rsIndex%>" value="Y" class="noborder" <% If Not IsValid Then %>disabled<% End If %> onclick="javascript:chkCheckAll('Active');"></td>
					<td>
					<% If Not IsValid Then %><a href="#" onclick="javascript:alert('<%=Replace(ErrMsg, "'", "\'")%>');"><% End If %>
					<% If Not IsValid Then %><font color="#FF0000"><% End If %><% If IsValid Then %><%=getadminRepImportLngStr("DtxtValid")%><% Else %><%=getadminRepImportLngStr("DtxtError")%></font><% End If %>
					<% If Not IsValid Then %></a><% End If %>
					</td>
				</tr>
				<% If IsValid Then allDisabled = ""
				Next %>
				<tr class="TblRep<% If Alter Then %>A<% End If %>Tbl">
					<td colspan="2">
					<p>
					<input type="checkbox" name="chkAllrsIndex" value="Y" id="chkAllRS" class="noborder" onclick="javascript:chkAll('rsIndex', this.checked);">
					<label for="chkAllRS"><%=getadminRepImportLngStr("DtxtAll")%></label></td>
					<td>
					&nbsp;</td>
					<td colspan="2">
					<p>
					<input type="checkbox" name="chkAllActive" value="Y" <%=allDisabled%> id="chkAllActive" class="noborder" onclick="javascript:chkAll('Active', this.checked);">
					<label for="chkAllActive"><%=getadminRepImportLngStr("DtxtAll")%></label></td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td bordercolor="#C8E7E8" colspan="2">
			<table border="0" cellpadding="0" cellspacing="1" width="100%" id="table3">
				<tr>
					<td width="75">
			<input type="submit" value="<%=getadminRepImportLngStr("DtxtSave")%>" name="btnSave" class="BtnRep"></td>
					<td>
			<hr size="1">
					</td>
					<td width="75">
					<p align="right">
			<input type="button" value="<%=getadminRepImportLngStr("DtxtCancel")%>" name="btnCancel" class="BtnRep" onclick="if(confirm('<%=getadminRepImportLngStr("DtxtConfCancel")%>'))window.close();"></td>
				</tr>
			</table>
			</td>
		</tr>
			<input type="hidden" name="UserType" value="<%=Request.QueryString("UserType")%>">
		<input type="hidden" name="FileName" value="<%=FileName%>">
		<input type="hidden" name="pop" value="Y">
		
		<% End If %>
</table>
</form>
<% End If %>
</body>
<% 
Function ValidateQuery(Query, QueryType)
	On Error Resume Next
	conn.execute(Query)
	If Err.Number <> 0 Then
		Select Case QueryType
			Case "Q"
				ErrMsg = "" & getadminRepImportLngStr("LtxtErrValQry") & ": \n"
			Case "V"
				ErrMsg = "" & getadminRepImportLngStr("LtxtErrValVar") & ": \n"
			Case "D"
				ErrMsg = "" & getadminRepImportLngStr("LtxtErrValDefVal") & ": \n"
		End Select
		ErrMsg = ErrMsg & Err.Description
		ValidateQuery = False
	Else
		ValidateQuery = True
	End If
End Function 

Sub loadRSVars()
	set rsVars = Server.CreateObject("ADODB.RecordSet")
	with rsVars.Fields
		.Append "varIndex", adInteger
		.Append "varVar", adVarWChar, 50
		.Append "varDataType", adVarChar, 10
		.Append "varQuery", adLongVarWChar, 1
		.Append "varMaxChar", adInteger
	end with
	rsVars.Open
	For each var in myVars
		rsVars.AddNew
		rsVars.Fields("varIndex")		= var.selectSingleNode("varIndex").text
		rsVars.Fields("varVar")			= var.selectSingleNode("varVar").text
		rsVars.Fields("varDataType")	= var.selectSingleNode("varDataType").text
		rsVars.Fields("varQuery")		= var.selectSingleNode("varQuery").text
		If var.selectSingleNode("varMaxChar").text <> "" Then
			rsVars.Fields("varMaxChar") = var.selectSingleNode("varMaxChar").text
		Else
			rsVars.Fields("varMaxChar") = -1
		End If
	next
End Sub
%>
</html>
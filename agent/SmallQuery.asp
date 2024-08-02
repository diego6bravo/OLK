<!--#include file="chkLogin.asp"-->
<!--#include file="lang/SmallQuery.asp" -->
<!--#include file="myHTMLEncode.asp"-->
<html <% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

varx = 0
set rs = Server.CreateObject("ADODB.RecordSet")
If userType = "V" Then
	SelDes = "0"
Else
	sql = "select SelDes from OLKCommon"
	set rs = conn.execute(sql)
	SelDes = rs(0)
End If

set rd = Server.CreateObject("ADODB.RecordSet")

Select Case Request("sType")
	Case "Act"
		TableID = "OCLG"
	Case "Item"
		TableID = "OITM"
	Case "Doc"
		TableID = "OINV"
	Case "Card"
		TableID = "OCRD"
	Case "Rec"
		TableID = "ORCT"
	Case "DocLine"
		TableID = "INV1"
	Case "Cnt"
		TableID = "OCPR"
	Case "Addr"
		TableID = "CRD1"
End Select

Select Case Request("source")
	Case "olkrep"
		Select Case Request("s") 
			Case "Q"
				sql = "select varName As Descr, varQuery As sqlQuery, varQueryField As SqlQueryField from OLKRSVars where rsIndex = " & Request("rsIndex") & " and varIndex = " & Request("varIndex")
			Case "F" 
				sql = 	"select IsNull(alterVarName, varName) as Descr, " & _
						"'select valValue As Value, IsNull(alterValText, valText) As Text " & _
						"from OLKRSVarsVals T0 " & _
						"left outer join OLKRSVarsValsAlterNames T1 on T1.rsIndex = T0.rsIndex and T1.varIndex = T0.varIndex and T1.valIndex = T0.valIndex and T1.LanID = " & Session("LanID") & " " & _
						"where T0.rsIndex = " & Request("rsIndex") & " and T0.varIndex = " & Request("varIndex") & "' SqlQuery, " & _
						"'Value' SqlQueryField " & _
						"from OLKRSVars T0 " & _
						"left outer join OLKRSVarsAlterNames T1 on T1.rsIndex = T0.rsIndex and T1.varIndex = T0.varIndex and T1.LanID = " & Session("LanID") & " " & _
						"where T0.rsIndex = " & Request("rsIndex") & " and T0.varIndex = " & Request("varIndex")
		End Select
	Case "customsearch"
		Select Case Request("s") 
			Case "Q"
				sql = "select Name As Descr, Query As sqlQuery, QueryField As SqlQueryField from OLKCustomSearchVars where ObjectCode = " & Request("ObjID") & " and ID = " & Request("ID") & " and VarID = " & Request("VarID")
			Case "F" 
				sql = 	"select IsNull(alterName, Name) as Descr, " & _
						"'select valValue As Value, IsNull(alterValText, valText) As Text " & _
						"from OLKCustomSearchVarsVals T0 " & _
						"left outer join OLKCustomSearchVarsValsAlterNames T1 on T1.ObjectCode = T0.ObjectCode and T1.ID = T0.ID and T1.VarID = T0.VarID and T1.ValID = T0.ValID and T1.LanID = " & Session("LanID") & " " & _
						"where T0.ObjectCode = " & Request("ObjID") & " and T0.ID = " & Request("ID") & " and T0.VarID = " & Request("VarID") & "' SqlQuery, " & _
						"'Value' SqlQueryField " & _
						"from OLKCustomSearchVars T0 " & _
						"left outer join OLKCustomSearchVarsAlterNames T1 on T1.ObjectCode = T0.ObjectCode and T1.ID = T0.ID and T1.VarID = T0.VarID and T1.LanID = " & Session("LanID") & " " & _
						"where T0.ObjectCode = " & Request("ObjID") & " and T0.ID = " & Request("ID") & " and T0.VarID = " & Request("VarID")
		End Select
	Case Else
		sql = "select Descr, SqlQuery, SqlQueryField " & _
			  "from cufd T0 " & _
			  "inner join OLKCUFD T1 on T1.TableID = T0.TableID and T1.FieldID = T0.FieldID " & _
			  "where T0.TableID = '" & TableID & "' and T0.FieldID = " & Request("FieldID")
End Select
set rs = conn.execute(sql)
sqlSmall = ""
If Request("source") <> "olkrep" Then
	LogNum = -1
	
	If Request("noLogNum") <> "Y" Then
		Select Case Request("sType")
			Case "Act"
				LogNum = Session("ActRetVal")
			Case "Item"
				LogNum = Session("ItmRetVal")
			Case "Card", "Cnt", "Addr"
				LogNum = Session("CrdRetVal")
			Case "Doc", "DocLine"
				LogNum = Session("RetVal")
			Case "Rec"
				LogNum = Session("PayRetVal")
		End Select
	End If
	
	sqlSmall = " declare @LogNum int set @LogNum = " & LogNum & " " & _
		"declare @dbName nvarchar(100) set @dbName = '" & Session("olkdb") & "' " & _
		"declare @branch int set @branch = " & Session("branch") & " declare @SlpCode int set @SlpCode = " & Session("vendid") & " "
	
	If Request("sType") <> "Item" Then
		sqlSmall = sqlSmall & "declare @CardCode nvarchar(15) set @CardCode = N'" & Session("username") & "' "
	End If
	
	If Request("sType") = "Doc" or Request("sType") = "DocLine" Then
		sqlSmall = sqlSmall & "declare @PriceList int set @PriceList = " & Session("PriceList") & " "
	End If
	
	If Request("sType") = "DocLine" Then
		sqlSmall = sqlSmall & "declare @ItemCode nvarchar(15) set @ItemCode = (select ItemCode from R3_ObsCommon..DOC1 where LogNum = @LogNum and LineNum = " & Request("LineNum") & ") " & _
							"declare @WhsCode nvarchar(8) set @WhsCode = (select WhsCode from R3_ObsCommon..DOC1 where LogNum = @LogNum and LineNum = " & Request("LineNum") & ") "
	End If

	sqlSmall = sqlSmall & rs("SqlQuery")
Else
	set rBase = Server.CreateObject("ADODB.RecordSet")
	sql = "select varIndex, varVar, varName, varDataType, varMaxChar  " & _
			"from OLKRSVars T0  " & _
			"where T0.rsIndex = " & Request("rsIndex") & " and varIndex in " & _
			"(select baseIndex from OLKRSVarsBase where rsIndex = T0.rsIndex and varIndex = " & Request("varIndex") & ") "
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
			sqlSmall = sqlSmall & "set @" & rBase("varVar") & " = '" & Request("var" & rBase("varIndex")) & "' "
		Else
			sqlSmall = sqlSmall & "set @" & rBase("VarVar") & " = " & Request("var" & rBase("varIndex")) & " "
		End If
	rBase.movenext
	loop
	'sqlSmall = sqlSmall & "declare @LanID int set @LanID = " & Session("LanID") & " "
	sqlSmall = sqlSmall & rs("SqlQuery")
End If
sqlSmall = "declare @LanID int set @LanID = " & Session("LanID") & " " & sqlSmall

rd.open QueryFunctions(sqlSmall), conn, 3, 1
If Not IsNull(rs("SqlQueryField")) Then SqlQueryField = rs("SqlQueryField") Else SqlQueryField = rd.Fields(0).Name
           %>
<link rel="stylesheet" type="text/css" href="design/<%=SelDes%>/style/stylePopUp.css">
<% If Session("Rtl") <> "" Then %><link rel="stylesheet" type="text/css" href="design/<%=SelDes%>/style/hebrewPop.css"><% End If %>
<title><%=rs("Descr")%></title>
</head>
<script language="javascript">
function setCartField(fieldVal)
{
<% If Request("MaxSize") = "" Then %>
	opener.setTimeStamp<%=Request("addCmd")%>('', fieldVal);
<% Else %>
	if (fieldVal.length > <%=Request("MaxSize")%>) fieldVal = fieldVal.substring(0, <%=Request("MaxSize")%>);
	opener.setTimeStamp<%=Request("addCmd")%>('', fieldVal);
<% End If %>
window.close();
}
</script>
<% Dim fSize()
Redim fSize(CInt(rd.Fields.Count-1))
For i = 0 to UBound(fSize)
	fSize(i) = 0
Next %>
<!--#include file="design/popvars.inc" -->
<body topmargin="0" leftmargin="0" onresize="<% If setCustTtl and userType = "C" Then %>javascript:setTtlBg(false);<% End If %>">
<table border="0" cellpadding="0" width="100%" id="table1">
	<% If tblCustTtl = "" Then %>
	<tr class="CSpecialTlt">
		<td id="tdMyTtl">&nbsp;<%=getSmallQueryLngStr("LtxtChangeVal")%></td>
	</tr>
	<% Else %>
	<%=Replace(Replace(tblCustTtl, "{txtTitle}", getSmallQueryLngStr("LtxtChangeVal")), "{AddPath}", "")%>
	<% End If %>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table2">
			<tr class="CSpecialTlt2">
			<% For each Field in rd.Fields %>
				<td>
				<p align="center"><%=Field.Name%>&nbsp;</td>
			<% next %>
			</tr>
			<% do while not rd.eof
			 %>
			<tr class="CSpecialTbl" style="cursor: hand" onclick="javascript:setCartField('<%=rd(CStr(SqlQueryField))%>')">
				<% 	For each Field in rd.Fields 
					If fSize(varx) < Len(Field) Then fSize(varx) = Len(Field)
					varx = varx + 1 %>
					<td>
					<p><% If Not IsNull(Field) Then %><%=Field%><% End If %>&nbsp;
					</td>
				<% 	next %>
			</tr>
		<% 	varx = 0
		  	rd.movenext
		  	loop %>
			<tr class="CSpecialTlt2">
			<% For i = 0 to UBound(fSize) %>
				<td>
				<% For j = 0 to (fSize(i)*2.8) %>&nbsp;<% Next %></td>
			<% next %>
			</tr>
		</table>
		</td>
	</tr>
</table>

<% If setCustTtl and userType = "C" Then %>
<script language="javascript" src="setTltBg.js.asp?custTtlBgL=<%=custTtlBgL%>&custTtlBgM=<%=custTtlBgM%>"></script>
<script language="javascript">setTtlBg(false);</script>
<% End If %>
</body>
<% conn.close
set rs = nothing
set rd = nothing %>
</html>

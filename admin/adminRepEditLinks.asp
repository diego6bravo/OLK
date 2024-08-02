<!--#include file="chkLogin.asp" -->
<!--#include file="lang/adminRepEditLinks.asp" -->
<!--#include file="myHTMLEncode.asp" -->
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<% 
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<html <% If Session("rtl") <> "" Then %>dir="rtl"<% End If %>>

<head>
<title><%=getadminRepEditLinksLngStr("LttlEditRepLnk")%> - <%=Request("colName")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<script language="javascript" src="general.js"></script>
<script type="text/javascript" src="scr/calendar.js"></script>
<script type="text/javascript" src="scr/lang/calendar-<%=Left(Session("myLng"), 2)%>.js"></script>
<script type="text/javascript" src="scr/calendar-setup.js"></script>
<link rel="stylesheet" type="text/css" href="style_cal.css">
<link rel="stylesheet" type="text/css" href="style/style_pop.css">
<style type="text/css">
.style1 {
	COLOR: #4783C5;
	FONT-SIZE: 10px;
	FONT-FAMILY: VERDANA;
	TEXT-DECORATION: none;
	font-weight: none;
}
.style2 {
	COLOR: #31659C;
	FONT-SIZE: 10px;
	FONT-FAMILY: VERDANA;
	TEXT-DECORATION: none;
	font-weight: bold;
}
.style3 {
	COLOR: #02E94D;
	FONT-SIZE: 10px;
	FONT-FAMILY: VERDANA;
	TEXT-DECORATION: none;
	font-weight: none;
}
</style>
</head>
<%
Dim ArrCol() %>
<!--#include file="repVars.inc" -->
<% 

sql = "select linkType, IsNull(linkObject,-1) linkObject, IsNull(linkObjectPocket,-1) linkObjectPocket, linkLink, linkLinkPocket, linkPopup, linkCat " & _
		"from OLKRSTotals " & _
		"where rsIndex = " & Request("rsIndex") & " and colName = N'" & saveHTMLDecode(Request("colName"), False) & "'"
set rs = conn.execute(sql)
If Request.Form("isSubmit") <> "Y" Then
	If Not rs.eof Then
		linkType = rs("linkType")
		linkObject = rs("linkObject")
		linkObjectPocket = rs("linkObjectPocket")
		linkLink = rs("linkLink")
		linkLinkPocket = rs("linkLinkPocket")
		linkPopup = rs("linkPopup")
		linkCat = rs("linkCat")
	Else
		linkType = "N"
		linkObject = "-1"
	End If
Else
	linkType = Request("linkType")
	linkObject = Request("linkObject")
	If linkObject = "" Then linkObject = "-1"
	linkObjectPocket = Request("linkObjectPocket")
	If linkObjectPocket = "" Then linkObjectPocket = "-1"
End If
%>
<script language="javascript">
function setTblSet()
{
	if (browserDetect() == 'msie')
	{
		tblSave.style.top = document.body.offsetHeight-31+document.body.scrollTop;
	}
	else if (browserDetect() == 'opera')
	{
		tblSave.style.top = document.body.offsetHeight-27+document.body.scrollTop;
	}
	else //firefox & others
	{
		tblSave.style.top = window.innerHeight-27+document.body.scrollTop;
	}
}
</script>
<body topmargin="0" leftmargin="0" onbeforeunload="opener.clearWin();" onload="setTblSet();" onscroll="setTblSet();">
<form method="POST" action="adminRepEditLinks.asp" name="frmLink" onsubmit="return valFrm();">
<table border="0" width="100%" id="table1" cellpadding="0">
	<tr class="TblGreenTlt">
		<td colspan="3"><%=getadminRepEditLinksLngStr("LttlEditRepLnk")%> - 
		 <%=Request("colName")%></td>
	</tr>
	<tr class="TblGreenNrm">
		<td colspan="3">&nbsp;</td>
	</tr>
	<tr>
		<td class="TblGreenTlt">
		<%=getadminRepEditLinksLngStr("DtxtType")%></td>
		<td colspan="2" class="TblGreenNrm">
		<font color="#4783C5" face="Verdana" size="1">
		<select size="1" name="linkType" id="linkType" onchange="if(document.frmLink.linkObject)document.frmLink.linkObject.selectedIndex=0;submit();">
		<option value="N"></option>
		<option <% If linkType = "A" Then %>selected<% End If %> value="A"><%=getadminRepEditLinksLngStr("LtxtAction")%></option>
		<option <% If linkType = "L" Then %>selected<% End If %> value="L"><%=getadminRepEditLinksLngStr("LtxtExtLnk")%></option>
		<option <% If linkType = "F" Then %>selected<% End If %> value="F"><%=getadminRepEditLinksLngStr("DtxtForm")%></option>
		<% If repEnObj Then %><option <% If linkType = "O" Then %>selected<% End If %> value="O"><%=getadminRepEditLinksLngStr("LtxtObject")%></option><% End If %>
		<option <% If linkType = "C" Then %>selected<% End If %> value="C"><%=getadminRepEditLinksLngStr("DtxtOp")%></option>
		<option <% If linkType = "R" Then %>selected<% End If %> value="R"><%=getadminRepEditLinksLngStr("LtxtRep")%></option>
		</select></font></td>
	</tr>
	<% Select Case linkType
		Case "N"
		Case "F" 
		LoadArrCols
		sql = "select SecID, SecName from OLKSections where SecType = 'U' and UserType = Case '" & Request("UserType") & "' When 'V' Then 'A' Else '" & Request("UserType") & "' End and Status = 'A' order by SecOrder"
		set rs = conn.execute(sql) %>
		<tr>
		<td class="TblGreenTlt">
		<%=getadminRepEditLinksLngStr("DtxtForm")%><% If Request("UserType") <> "C" Then %> (<%=getadminRepEditLinksLngStr("DtxtAgent")%>)<% End If %></td>
		<td colspan="2" class="TblGreenNrm">
		<font color="#4783C5" face="Verdana" size="1">
		<select size="1" name="linkObject" id="linkObject" onchange="document.frmLink.linkLink.value = '';document.getElementById('linkLinkData').style.display = this.selectedIndex<=0?'none':'';">
		<option value="">--&gt; <%=getadminRepEditLinksLngStr("LtxtSelObj")%> &lt;--</option>
		<% do while not rs.eof %>
		<option <% If CStr(rs(0)) = CStr(linkObject) Then %>selected<% End If %> value="<%=rs(0)%>"><%=rs(1)%></option>
		<% rs.movenext
		loop %>
		</select>
		</font>
		</td>
		</tr>
		<tbody id="linkLinkData" <% If linkObject = "-1" Then %>style="display: none;"<% End If %>>
		<tr class="TblGreenNrm">
			<td colspan="3">
			<textarea dir="ltr" rows="3" class="input" name="linkLink" style="width: 100%"><%=myHTMLEncode(linkLink)%></textarea>
			</td>
		</tr>
		<tr>
			<td valign="top" class="TblGreenTlt"><%=getadminRepEditLinksLngStr("DtxtVariables")%>
			</td>
			<td colspan="2" class="TblGreenNrm">
			<font color="#4783C5" face="Verdana" size="1">
			<select size="8" name="linkLinkVars" onclick="javascript:if(this.value!=null&&this.value!='')document.frmLink.linkLink.value+='{' + this.value + '}';">
			<% For i = 0 to UBound(ArrCol) %>
			<option value="<%=myHTMLEncode(ArrCol(i)(1))%>"><%=myHTMLEncode(ArrCol(i)(0))%></option>
			<% Next %>
			</select></font></td>
		</tr>
		</tbody>
		<% If Request("UserType") <> "C" Then
		sql = "select SecID, SecName from OLKSections where SecType = 'U' and UserType = 'P' and Status = 'A' order by SecOrder"
		set rs = conn.execute(sql) %>
		<tr>
		<td class="TblGreenTlt">
		<%=getadminRepEditLinksLngStr("DtxtForm")%> (<%=getadminRepEditLinksLngStr("DtxtPocket")%>)</td>
		<td colspan="2" class="TblGreenNrm">
		<font color="#4783C5" face="Verdana" size="1">
		<select size="1" name="linkObjectPocket" id="linkObjectPocket" onchange="document.frmLink.linkLinkPocket.value = '';document.getElementById('linkLinkPocketData').style.display = this.selectedIndex<=0?'none':'';">
		<option value="">--&gt; <%=getadminRepEditLinksLngStr("LtxtSelObj")%> &lt;--</option>
		<% do while not rs.eof %>
		<option <% If CStr(rs(0)) = CStr(linkObjectPocket) Then %>selected<% End If %> value="<%=rs(0)%>"><%=rs(1)%></option>
		<% rs.movenext
		loop %>
		</select>
		</font>
		</td>
		</tr>
		<tbody id="linkLinkPocketData" <% If linkObjectPocket = "-1" Then %>style="display: none;"<% End If %>>
		<tr class="TblGreenNrm">
			<td colspan="3">
			<textarea dir="ltr" rows="3" class="input" name="linkLinkPocket" style="width: 100%"><%=myHTMLEncode(linkLinkPocket)%></textarea>
			</td>
		</tr>
		<tr>
			<td valign="top" class="TblGreenTlt"><%=getadminRepEditLinksLngStr("DtxtVariables")%>
			</td>
			<td colspan="2" class="TblGreenNrm">
			<font color="#4783C5" face="Verdana" size="1">
			<select size="8" name="linkLinkPocketVars" onclick="javascript:if(this.value!=null&&this.value!='')document.frmLink.linkLinkPocket.value+='{' + this.value + '}';">
			<% For i = 0 to UBound(ArrCol) %>
			<option value="<%=myHTMLEncode(ArrCol(i)(1))%>"><%=myHTMLEncode(ArrCol(i)(0))%></option>
			<% Next %>
			</select></font></td>
		</tr>
		</tbody>
		<tr>
			<td style="height: 24px;"></td>
		</tr>
		<% End If %>
<%		Case "L"
			LoadArrCols %>
			<tr>
				<td colspan="3" class="TblGreenTlt">
				<p align="center"><%=getadminRepEditLinksLngStr("LttlLinkBuilder")%></td>
			</tr>
			<tr class="TblGreenNrm">
				<td colspan="3">
				<textarea dir="ltr" rows="3" class="input" name="linkLink" style="width: 100%"><%=myHTMLEncode(linkLink)%></textarea>
				</td>
			</tr>
			<tr>
				<td valign="top" class="TblGreenTlt"><%=getadminRepEditLinksLngStr("DtxtVariables")%>
				</td>
				<td colspan="2" class="TblGreenNrm">
				<font color="#4783C5" face="Verdana" size="1">
				<select size="8" name="linkLinkVars" onclick="javascript:if(this.value!=null&&this.value!='')document.frmLink.linkLink.value+='{' + this.value + '}';">
				<% For i = 0 to UBound(ArrCol) %>
				<option value="<%=myHTMLEncode(ArrCol(i)(1))%>"><%=myHTMLEncode(ArrCol(i)(0))%></option>
				<% Next %>
				</select></font></td>
			</tr>
	<% Case Else %>
	<tr>
		<td class="TblGreenTlt">
		<% 
		Select Case linkType 
			Case "R" %><%=getadminRepEditLinksLngStr("LtxtRep")%><% 
			Case "O" %><%=getadminRepEditLinksLngStr("LtxtObject")%><% 
			Case "C" %><%=getadminRepEditLinksLngStr("DtxtOp")%><%
		End Select %></td>
		<td colspan="2" class="TblGreenNrm">
		<font color="#4783C5" face="Verdana" size="1">
		<select size="1" name="linkObject" id="linkObject" onchange="submit();">
		<option value="">--&gt; <%=getadminRepEditLinksLngStr("LtxtSelObj")%> &lt;--</option>
		<% Select Case linkType 
			Case "F" 
			%>
		<%	Case "A" %>
		<optgroup label="<%=getadminRepEditLinksLngStr("DtxtSAP")%>">
		<option <% If CStr(linkObject) = "1" Then %>selected<% End If %> value="1"><%=getadminRepEditLinksLngStr("LtxtConvQuoteOrder")%></option>
		<option <% If CStr(linkObject) = "10" Then %>selected<% End If %> value="10"><%=getadminRepEditLinksLngStr("LtxtConvOrderDel")%></option>
		<option <% If CStr(linkObject) = "7" Then %>selected<% End If %> value="7"><%=getadminRepEditLinksLngStr("LtxtConvOrderInv")%></option>
		<option <% If CStr(linkObject) = "8" Then %>selected<% End If %> value="8"><%=getadminRepEditLinksLngStr("LtxtAprovPurOrdr")%></option>
		<option <% If CStr(linkObject) = "0" Then %>selected<% End If %> value="0"><%=getadminRepEditLinksLngStr("LtxtAprovOrder")%></option>
		<option <% If CStr(linkObject) = "2" Then %>selected<% End If %> value="2"><%=getadminRepEditLinksLngStr("LtxtCloseObj")%></option>
		<option <% If CStr(linkObject) = "3" Then %>selected<% End If %> value="3"><%=getadminRepEditLinksLngStr("LtxtCancelObj")%></option>
		<option <% If CStr(linkObject) = "6" Then %>selected<% End If %> value="6"><%=getadminRepEditLinksLngStr("LtxtRemObj")%></option>
		</optgroup>
		<optgroup label="<%=getadminRepEditLinksLngStr("DtxtOLK")%>">
		<option <% If CStr(linkObject) = "4" Then %>selected<% End If %> value="4"><%=getadminRepEditLinksLngStr("LtxtAddItmCart")%></option>
		<option <% If CStr(linkObject) = "5" Then %>selected<% End If %> value="5"><%=getadminRepEditLinksLngStr("LtxtAddItmWish")%></option>
		<option <% If CStr(linkObject) = "9" Then %>selected<% End If %> value="9"><%=getadminRepEditLinksLngStr("LtxtDupDoc")%></option>
		</optgroup>
		<%	Case "O" %>
		<optgroup label="<%=getadminRepEditLinksLngStr("LtxtGeneral")%>">
		<% myObjs = "33, " & getadminRepEditLinksLngStr("DtxtActivity") & "{S}"
		If Request("UserType") = "V" Then myObjs = myObjs & "2, " & getadminRepEditLinksLngStr("DtxtClient") & "{S}" 
		myObjs = myObjs & _
					"4, " & getadminRepEditLinksLngStr("DtxtItem") & "{S}" & _
					"97, " & getadminRepEditLinksLngStr("DtxtSO") & "{S}" & _
					", " & getadminRepEditLinksLngStr("LtxtSale") & "{S}" & _
					"23, " & getadminRepEditLinksLngStr("DtxtQuote") & "{S}" & _
					"17, " & getadminRepEditLinksLngStr("DtxtSalesOrder") & "{S}" & _
					"15, " & getadminRepEditLinksLngStr("DtxtDelivery") & "{S}" & _
					"16, " & getadminRepEditLinksLngStr("LtxtReturn") & "{S}" & _
					"13, " & getadminRepEditLinksLngStr("DtxtInvoice") & "{S}" & _
					"14, " & getadminRepEditLinksLngStr("LtxtCredNote") & "{S}" & _
					", " & getadminRepEditLinksLngStr("LtxtPur") & "{S}" & _
					"540000006, " & getadminRepEditLinksLngStr("DtxtPurQuote") & "{S}" & _
					"22, " & getadminRepEditLinksLngStr("LtxtPurOrdr") & "{S}" & _
					"20, " & getadminRepEditLinksLngStr("LtxtGoodRecPO") & "{S}" & _
					"21, " & getadminRepEditLinksLngStr("LtxtPurReturn") & "{S}" & _
					"18, " & getadminRepEditLinksLngStr("LtxtPurInv") & "{S}" & _
					"19, " & getadminRepEditLinksLngStr("LtxtCredMemPO") & "{S}" & _
					", " & getadminRepEditLinksLngStr("LtxtBanks") & "{S}" & _
					"24, " & getadminRepEditLinksLngStr("DtxtReceipt") & "{S}" & _
					"46, " & getadminRepEditLinksLngStr("LtxtProvPay") & "{S}" & _
					", " & getadminRepEditLinksLngStr("DtxtDraft") & "{S}" & _
					"112, " & getadminRepEditLinksLngStr("LtxtSale") & " / " & getadminRepEditLinksLngStr("LtxtPur") & "{S}" & _
					"140, " & getadminRepEditLinksLngStr("DtxtReceipt") & " / " & getadminRepEditLinksLngStr("LtxtProvPay") & "{S}" & _
					", OLK{S}" & _
					"-9, " & getadminRepEditLinksLngStr("DtxtLogNum") & " - " & getadminRepEditLinksLngStr("DtxtActivity") & "{S}" & _
					"-7, " & getadminRepEditLinksLngStr("DtxtLogNum") & " - " & getadminRepEditLinksLngStr("DtxtClient") & "{S}" & _
					"-8, " & getadminRepEditLinksLngStr("DtxtLogNum") & " - " & getadminRepEditLinksLngStr("DtxtItem") & "{S}" & _
					"-4, " & getadminRepEditLinksLngStr("DtxtLogNum") & " - " & getadminRepEditLinksLngStr("DtxtComDocs") & "{S}" & _
					"-6, " & getadminRepEditLinksLngStr("DtxtLogNum") & " - " & getadminRepEditLinksLngStr("DtxtReceipt") & "{S}" & _
					"-11, " & getadminRepEditLinksLngStr("DtxtLogNum") & " - " & getadminRepEditLinksLngStr("DtxtSO") & "{S}" & _
					"-12, " & getadminRepEditLinksLngStr("DtxtLogNum") & " - " & getadminRepEditLinksLngStr("LtxtFlowViewControl") & "{S}" & _
					"-5, " & getadminRepEditLinksLngStr("LtxtDynObj") & "{S}" & _
					"-10, " & getadminRepEditLinksLngStr("LttlAcctBal") & ""
		arrObj = Split(myObjs, "{S}")
		For i = 0 to UBound(arrObj)
		objItm = Split(arrObj(i), ", ")
		If objItm(0) <> "" Then %>
		<option <% If CStr(linkObject) = CStr(objItm(0)) Then %>selected<% End If %> value="<%=objItm(0)%>"><%=objItm(1)%></option>
		<% Else 
		Response.Write "</optgroup><optgroup label=""" & objItm(1) & """>"
		End If %>
		<% Next %>
		</optgroup>
		<% Case "R"
		lRG = -1
		sql = "declare @rsIndex int set @rsIndex = " & Request("rsIndex") & " " & _
		"select T0.rsIndex, T0.rsName, T1.rgIndex, T1.rgName " & _
		"from OLKRS T0 " & _
		"inner join OLKRG T1 on T1.rgIndex = T0.rgIndex " & _
		"where T1.UserType = '" & Request("UserType") & "' " & _
		"and T0.rsIndex <> @rsIndex and T0.Active = 'Y' " & _
		"order by T1.rgName, T0.rsName "
		set rs = conn.execute(sql)
		If rs.Eof Then ignoreGrp = True Else ignoreGrp = False
		do while not rs.eof
		If lRG <> CInt(rs("rgIndex")) Then
			If lRG <> -1 and lRG <> CInt(rs("rgIndex")) Then %>
			</optgroup>
			<% End If %>
		<optgroup label="<%=myHTMLEncode(rs("rgName"))%>">
		<% lRG = CInt(rs("rgIndex"))
		End If %>
		<option <% If CStr(linkObject) = CStr(rs(0)) Then %>selected<% End If %> value="<%=rs(0)%>"><%=myHTMLEncode(rs(1))%></option>
		<% rs.movenext
		   loop %>
		<% If Not ignoreGrp Then %></optgroup><% End If %>
		<% Case "C"
		lRG = -1
		sql = "select T0.ID, T0.Name, T1.ID GroupID, T1.Name GroupName " & _
		"from OLKOps T0 " & _
		"inner join OLKOpsGrps T1 on T1.ID = T0.GroupID " & _
		"where T0.Status = 'A' " & _
		"order by 4, 2 "
		set rs = conn.execute(sql)
		If rs.Eof Then ignoreGrp = True Else ignoreGrp = False
		do while not rs.eof
		If lRG <> CInt(rs("GroupID")) Then
			If lRG <> -1 and lRG <> CInt(rs("GroupID")) Then %>
			</optgroup>
			<% End If %>
		<optgroup label="<%=myHTMLEncode(rs("GroupName"))%>">
		<% lRG = CInt(rs("GroupID"))
		End If %>
		<option <% If CStr(linkObject) = CStr(rs(0)) Then %>selected<% End If %> value="<%=rs(0)%>"><%=myHTMLEncode(rs(1))%></option>
		<% rs.movenext
		   loop %>
		<% If Not ignoreGrp Then %></optgroup><% End If %>
		<% End Select%>
		</select></font></td>
	</tr>
	<% If linkObject <> "" and linkObject <> "-1" Then
		Select Case linkType 
			Case "A"
				Select Case CInt(linkObject)
					Case 0, 8
						sql = "select 'Entry' varVar, 'int' varDataType, 'Y' varNotNull"
					Case 1, 7, 10
						sql = "select 'Entry' varVar, 'int' varDataType, 'Y' varNotNull " & _
								"union select 'Series' varVar, 'int' varDataType, 'Y' varNotNull "
					Case 2, 3, 6, 9
						sql = "select 'ObjectCode' varVar, 'int' varDataType, 'Y' varNotNull " & _
								"union select 'Entry' varVar, 'int' varDataType, 'Y' varNotNull"
					Case 4 'Add Item
						sql = "select 'ItemCode' varVar, 'nvarchar' varDataType, 'Y' varNotNull " & _
								"union select 'Quantity' varVar, 'float' varDataType, 'N' varNotNull " & _
								"union select 'Unit' varVar, 'int' varDataType, 'N' varNotNull " & _
								"union select 'Price' varVar, 'float' varDataType, 'N' varNotNull " & _
								"union select 'Locked' varVar, 'nvarchar' varDataType, 'N' varNotNull " & _
								"union select 'WhsCode' varVar, 'nvarchar' varDataType, 'N' varNotNull "
					Case 5 
						sql = "select 'ItemCode' varVar, 'nvarchar' varDataType, 'Y' varNotNull "
				End Select
			Case "O"
				Select Case CLng(linkObject)
					Case 2
						sql = "select 'CardCode' varVar, 'nvarchar' varDataType, 'Y' varNotNull"
					Case 4
						sql = "select 'ItemCode' varVar, 'nvarchar' varDataType, 'Y' varNotNull"
					Case 13, 17, 23, 15, 16, 13, 14, 22, 20, 21, 18, 19, 46, 24, 112, 140, 540000006
						sql = 	"select 'DocEntry' varVar, 'int' varDataType, 'Y' varNotNull union " & _
								"select 'LineNum' varVar, 'int' varDataType, 'N' varNotNull"
					Case 33
						sql = "select 'ClgCode' varVar, 'int' varDataType, 'Y' varNotNull"
					Case 97
						sql = "select 'OpprId' varVar, 'int' varDataType, 'Y' varNotNull"
					Case -2
						sql = 	"select 'CardCode' varVar, 'nvarchar' varDataType, 'Y' varNotNull union " & _
								"select 'dtFrom' varVar, 'datetime' varDataType, 'N' varNotNull union " & _
								"select 'dtTo' varVar, 'datetime' varDataType, 'N' varNotNull"
					Case -4, -6, -7, -8, -9, -11
						sql = 	"select 1 Ordr, 'Entry' varVar, 'int' varDataType, 'Y' varNotNull "
						If linkObject = "-4" Then sql = sql & " union select 2 Ordr, 'LineNum' varVar, 'int' varDataType, 'N' varNotNull"
					Case -12
						sql = 	"select 1 Ordr, 'ID' varVar, 'int' varDataType, 'Y' varNotNull "
					Case -5
						sql = 	"select 'ObjectCode' varVar, 'int' varDataType, 'Y' varNotNull union " & _
								"select 'DocNum' varVar, 'int' varDataType, 'Y' varNotNull  union " & _
								"select 'LineNum' varVar, 'int' varDataType, 'N' varNotNull "
					Case -10
						sql =	"select 'CardCode' varVar, 'nvarchar' varDataType, 'Y' varNotNull"
				End Select
			Case "R" 
				sql = 	"select T0.varVar, T0.varDataType, IsNull(T1.valBy,'F') valBy, T1.valValue, T1.valValDat, T0.varNotNull " & _
						"from OLKRSVars T0 " & _
						"left outer join OLKRSLinksVars T1 on T1.rsIndex = " & Request("rsIndex") & " and T1.colName = N'" & saveHTMLDecode(Request("colName"), False) & "' and varID = T0.varVar " & _
						"where T0.rsIndex = " & linkObject & " order by Ordr"
			Case "C"
				sql = "select T0.varVar, T0.varDataType, IsNull(T2.valBy,'F') valBy, T2.valValue, T2.valValDat, 'Y' varNotNull " & _  
				"from ( " & _  
				"	select 'CardCode' varVar, 'nvarchar' varDataType, 'Y' varNotNull, 2 ObjectID " & _  
				"	union all " & _  
				"	select 'ItemCode' varVar, 'nvarchar' varDataType, 'Y' varNotNull, 4 ObjectID " & _  
				"	union all " & _  
				"	select * from " & _  
				"	(			 " & _  
				"		select 'DocEntry' varVar, 'int' varDataType, 'Y' varNotNull " & _  
				"	) X0 " & _  
				"	cross join ( " & _  
				"		select 13 ObjID " & _  
				"		union select 17 " & _  
				"		union select 23 " & _  
				"		union select 15 " & _  
				"		union select 16 " & _  
				"		union select 13 " & _  
				"		union select 14 " & _  
				"		union select 22 " & _  
				"		union select 20 " & _  
				"		union select 21 " & _  
				"		union select 18 " & _  
				"		union select 19 " & _  
				"		union select 46 " & _  
				"		union select 24 " & _  
				"		union select 112 " & _  
				"		union select 140 " & _  
				"		union select 540000006 " & _  
				"	) X1 " & _  
				"	union all select 'ClgCode' varVar, 'int' varDataType, 'Y' varNotNull, 33 ObjectID " & _  
				"	union all select 'OpprId' varVar, 'int' varDataType, 'Y' varNotNull, 97 ObjectID " & _  
				") T0 " & _  
				"inner join OLKOps T1 on T1.ID = " & linkObject & " and T1.ObjectID = T0.ObjectID " & _
				"left outer join OLKRSLinksVars T2 on T2.rsIndex = " & Request("rsIndex") & " and T2.colName = N'" & saveHTMLDecode(Request("colName"), False) & "' and T2.varID = T0.varVar "
		End Select
		
		If linkType = "O" or linkType = "A" Then
			valBy = "IsNull(T1.valBy,'F')"
			If linkType = "A" and CLng(linkObject) = 1 Then
				valBy = "IsNull(T1.valBy,Case T0.varVar When 'Series' Then 'V' Else 'F' End)"
			End If
			sql = 	"select T0.varVar, T0.varDataType, " & valBy & " valBy, valValue, T1.valValDat, T0.varNotNull " & _
					"from (" & sql  & ") T0 " & _
					"left outer join OLKRSLinksVars T1 on T1.rsIndex = " & Request("rsIndex") & " and T1.colName = N'" & saveHTMLDecode(Request("colName"), False) & "' and T1.varId = T0.varVar "
		End If
		
		rs.close
		rs.open sql, conn, 3, 1
	 %>
	<% If linkType = "R" or linkType = "O" and (linkObject = -10) Then %>
	<tr>
		<td class="TblGreenTlt">
		&nbsp;</td>
		<td colspan="2" class="TblGreenNrm">
		<input type="checkbox" name="Popup" class="OptionButton" style="background:background-image" value="Y" id="Popup" <% If linkPopup = "Y" Then %>checked<% End If %>><label for="Popup"><%=getadminRepEditLinksLngStr("LtxtNewWin")%></label></td>
	</tr>
	<% End If %>
	<% If linkType = "O" and (linkObject = -4 or linkObject >= 13 and linkObject <= 23) Then %>
	<tr>
		<td class="TblGreenTlt">
		&nbsp;</td>
		<td colspan="2" class="TblGreenNrm">
		<input class="OptionButton" <% If linkCat = "Y" Then %>checked<% End If %> style="height: 20px; background:background-image" type="checkbox" value="Y" name="linkCat" id="linkCat"><label for="linkCat"><%=getadminRepEditLinksLngStr("DtxtCat")%></label></td>
	</tr>
	<% End If %>
	<tr>
		<td class="TblGreenTlt" colspan="3">
		<%=getadminRepEditLinksLngStr("DtxtVariables")%></td>
	</tr>
	<% if not rs.eof then
	LoadArrCols	
	do while not rs.eof %>
	<tr>
		<td dir="ltr" align="<% If Session("rtl") = "" Then %>left<% Else %>right<% End If %>" class="TblGreenTlt">
		@<%=rs("varVar")%><% If rs("varNotNull") = "Y" Then %><font color="red">*</font><% End If %></td>
		<td class="TblGreenNrm">
		<input type="radio" class="OptionButton" style="background:background-image" value="F" <% If rs("valBy") = "F" Then %>checked<% End If %> name="valBy<%=rs("varVar")%>" id="rdFld<%=rs("varVar")%>" onclick="changeValBy('<%=rs("varVar")%>','F');"><label for="rdFld<%=rs("varVar")%>"><%=getadminRepEditLinksLngStr("DtxtField")%></label><input class="OptionButton" style="background:background-image" type="radio" <% If rs("valBy") = "V" Then %>checked<% End If %> name="valBy<%=rs("varVar")%>" value="V" id="rdVal<%=rs("varVar")%>" onclick="changeValBy('<%=rs("varVar")%>','V');"><label for="rdVal<%=rs("varVar")%>"><%=getadminRepEditLinksLngStr("DtxtValue")%></label></td>
		<td class="TblGreenNrm">
		<font color="#4783C5" face="Verdana" size="1">
		<table border="0" id="tblValDat<%=rs("varVar")%>" cellspacing="0" cellpadding="0" style="<% If rs("valBy") = "V" and rs("varDataType") <> "datetime" or rs("valBy") = "F" Then %>;display: none<% End If %>">
			<tr>
				<td><img border="0" src="images/cal.gif" id="btnValDatImg<%=rs("varVar")%>" width="16" height="16" style="float:left;padding-left:1px;padding-top:1px"></td>
				<td>
				<input type="text" readonly name="colValDat<%=rs("varVar")%>" size="12" value="<%=FormatDate(rs("valValDat"), False)%>" onclick="btnValDatImg<%=rs("varVar")%>.click()"></td>
				<td><img border="0" src="images/remove.gif" style="cursor: hand" onclick="javascript:document.frmLink.colValDat<%=rs("varVar")%>.value='';"></td>
			</tr>
		</table>
		<% If linkType = "A" and (linkObject = 1 or linkObject = 7 or linkObject = 10) and rs("varVar") = "Series" Then %>
				<select size="1" name="valValueVSeries" id="valValueVSeries" class="input" style="<% If rs("valBy") = "F" Then %>display: none<% End If %>">
				<option value=""><%=getadminRepEditLinksLngStr("LtxtSelSeries")%></option>
					<% 
					If rs("valBy") = "V" and not IsNull(rs("valValue")) Then Series = CInt(rs("valValue")) Else Series = -1
					set rVal = Server.CreateObject("ADODB.RecordSet")
					seriesObj = 17
					Select Case linkObject 
						Case 7 
							seriesObj = 13
						Case 10
							seriesObj = 15
					End Select
					GetQuery rVal, 4, seriesObj, null
						do While NOT rVal.EOF %>
						<option <% If CInt(rVal("Series")) = Series Then %>selected<%end if%> value="<%=rVal("Series")%>"><%=myHTMLEncode(rVal("SeriesName"))%></option>
					<% rVal.MoveNext
					loop %>
				</select>
		<% ElseIf linkType = "A" and (linkObject = 2 or linkObject = 3 or linkObject = 6 or linkObject = 9) and rs("varVar") = "ObjectCode" Then %>
				<select size="1" name="valValueVObjectCode" id="valValueVObjectCode" class="input" style="<% If rs("valBy") = "F" Then %>display: none<% End If %>">
				<option value=""></option>
				<% Select Case linkObject
					Case 2 %>
					<option value="23"><%=getadminRepEditLinksLngStr("DtxtQuote")%></option>
					<option value="17"><%=getadminRepEditLinksLngStr("DtxtSalesOrder")%></option>
					<option value="15"><%=getadminRepEditLinksLngStr("DtxtDelivery")%></option>
					<option value="16"><%=getadminRepEditLinksLngStr("LtxtReturn")%></option>
					<option value="22"><%=getadminRepEditLinksLngStr("LtxtPurOrdr")%></option>
					<option value="20"><%=getadminRepEditLinksLngStr("LtxtGoodRecPO")%></option>
					<option value="191"><%=getadminRepEditLinksLngStr("DtxtServiceCall")%></option>
				<%	Case 3 %>
					<option value="23"><%=getadminRepEditLinksLngStr("DtxtQuote")%></option>
					<option value="17"><%=getadminRepEditLinksLngStr("DtxtSalesOrder")%></option>
					<option value="22"><%=getadminRepEditLinksLngStr("LtxtPurOrdr")%></option>
					<option value="24"><%=getadminRepEditLinksLngStr("DtxtReceipt")%></option>
					<option value="46"><%=getadminRepEditLinksLngStr("LtxtProvPay")%></option>
				<%	Case 6 %>
					<option value="2"><%=getadminRepEditLinksLngStr("DtxtClient")%></option>
					<option value="4"><%=getadminRepEditLinksLngStr("DtxtItem")%></option>
				<%	Case 9, 10 %>
					<option value="23"><%=getadminRepEditLinksLngStr("DtxtQuote")%></option>
					<option value="17"><%=getadminRepEditLinksLngStr("DtxtSalesOrder")%></option>
					<option value="15"><%=getadminRepEditLinksLngStr("DtxtDelivery")%></option>
					<option value="16"><%=getadminRepEditLinksLngStr("LtxtReturn")%></option>
					<option value="22"><%=getadminRepEditLinksLngStr("LtxtPurOrdr")%></option>
					<option value="20"><%=getadminRepEditLinksLngStr("LtxtGoodRecPO")%></option>
					<option value="22"><%=getadminRepEditLinksLngStr("LtxtPurOrdr")%></option>
				<%	End Select %>
				</select>
		<% ElseIf linkType = "A" and linkObject = 4 and (rs("varVar") = "Locked" or rs("varVar") = "Unit" or rs("varVar") = "WhsCode") Then
		Select Case rs("varVar")
		Case "Locked" %>
		<input type="checkbox" name="valValueVLocked" class="noborder" id="valValueVLocked" style="<% If rs("valBy") = "F" Then %>display: none<% End If %>" <% If rs("valBy") = "V" and rs("valValue") = "Y" Then %>checked<% End If %> value="Y">
		<% Case "Unit" %>
		<select size="1" name="valValueVUnit" id="valValueVUnit" style="<% If rs("valBy") = "F" Then %>display: none<% End If %>">
		<option></option>
		<option <% If rs("valBy") = "V" and rs("valValue") = "1" Then %>selected<% End If %> value="1"><%=getadminRepEditLinksLngStr("DtxtUnit")%></option>
		<option <% If rs("valBy") = "V" and rs("valValue") = "2" Then %>selected<% End If %> value="2"><%=getadminRepEditLinksLngStr("DtxtSalUnit")%></option>
		<option <% If rs("valBy") = "V" and rs("valValue") = "3" Then %>selected<% End If %> value="3"><%=getadminRepEditLinksLngStr("DtxtPackUnit")%></option>
		</select>
		<% Case "WhsCode" 
		set rw = Server.CreateObject("ADODB.RecordSet")
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKGetWarehouses" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		rw.open cmd, , 3, 1
 %>
		<select size="1" name="valValueVWhsCode" id="valValueVWhsCode" style="<% If rs("valBy") = "F" Then %>display: none<% End If %>">
		<option></option>
		<% do while not rw.eof %>
		<option <% If rs("valBy") = "V" and rs("valValue") = rw("WhsCode") Then %>selected<% End If %> value="<%=Server.HTMLEncode(rw("WhsCode"))%>"><%=Server.HTMLEncode(rw("WhsName"))%></option>
		<% rw.movenext
		loop %>
		</select>
		<% End Select
		Else %>
		<input style="<% If rs("valBy") = "F" or rs("varDataType") = "datetime" Then %>display: none<% End If %>" type="text" name="valValueV<%=rs("varVar")%>" id="valValueV<%=rs("varVar")%>" size="25" value="<% If rs("valBy") = "V" and not IsNull(rs("valValue")) Then Response.Write Server.HTMLEncode(rs("valValue"))%>" onchange="valThis(this,'<%=rs("varVar")%>');"><% End If %><select style="<% If rs("valBy") = "V" Then %>display: none<% End If %>" size="1" name="valValueF<%=rs("varVar")%>" id="valValueF<%=rs("varVar")%>">
		<option></option>
		<% For i = 0 to UBound(ArrCol)
		If 	ArrCol(i)(2) = "T" and (rs("varDataType") = "nvarchar" or linkObject = "-5") or _
			ArrCol(i)(2) = "D" and rs("varDataType") = "datetime" or _
			ArrCol(i)(2) = "N" and (rs("varDataType") = "float" or rs("varDataType") = "numeric" or rs("varDataType") = "int") Then %>
		<option <% If rs("valBy") = "F" Then If myHTMLEncode(rs("valValue")) = ArrCol(i)(1) Then Response.Write "selected" %> value="<%=ArrCol(i)(1)%>"><%=ArrCol(i)(0)%></option>
		<% End If
		   Next %>
		</select></font></td>
		<input type="hidden" name="varDataType<%=rs("varVar")%>" id="varDataType<%=rs("varVar")%>" value="<%=rs("varDataType")%>">
		<input type="hidden" name="varVar" value="<%=rs("varVar")%>">
		<input type="hidden" name="varNotNull" value="<%=rs("varNotNull")%>">
		<% If rs("varDataType") = "datetime" Then %>
		<script language="javascript">Calendar.setup({
		    inputField     :    "colValDat<%=rs("varVar")%>",     // id of the input field
		    ifFormat       :    "<%=GetCalendarFormatString%>",      // format of the input field
		    button         :    "btnValDatImg<%=rs("varVar")%>",  // trigger for the calendar (button ID)
		    align          :    "Bl",           // alignment (defaults to "Bl")
		    singleClick    :    true
		});</script>
		<% End If %>
	</tr>
	<% rs.movenext
	loop %>
	<% else %>
	<tr>
		<td align="right" colspan="3" class="TblGreenTlt">
		<p align="center"><%=getadminRepEditLinksLngStr("LtxtNoVars")%></td>
	</tr>
	<% End If %>
	<% End If %>
	<% End Select %>
</table>
		<table cellpadding="0" border="0" id="tblSave" style="width: 100%; position: absolute; z-index: 1; background-color: white;">
			<tr>
				<td style="width: 75px">
				<input type="submit" value="<%=getadminRepEditLinksLngStr("DtxtSave")%>" name="btnSave" class="OlkBtn"></td>
				<td><hr size="1"></td>
				<td style="width: 75px">
				<input type="button" value="<%=getadminRepEditLinksLngStr("DtxtCancel")%>" name="btnCancel" onclick="window.close();" class="OlkBtn"></td>
			</tr>
		</table>
	<input type="hidden" name="colName" value="<%=Request("colName")%>">
	<input type="hidden" name="rsIndex" value="<%=Request("rsIndex")%>">
	<input type="hidden" name="pop" value="<%=Request("pop")%>">
	<input type="hidden" name="UserType" value="<%=Request("UserType")%>">
	<input type="hidden" name="id" value="<%=Request("id")%>">
	<input type="hidden" name="isSubmit" value="Y">
	</form>
<script language="javascript">
function changeValBy(varVar, by)
{
	switch (by)
	{
		case 'F':
			document.getElementById('valValueV' + varVar).style.display='none';
			document.getElementById('tblValDat' + varVar).style.display='none';
			document.getElementById('valValueF' + varVar).style.display='';
			break;
		case 'V':
			if (document.getElementById('varDataType' + varVar).value != 'datetime')
			{
				document.getElementById('valValueV' + varVar).style.display='';
			}
			else
			{
				document.getElementById('tblValDat' + varVar).style.display='';
			}
			document.getElementById('valValueF' + varVar).style.display='none';
			break;
	}
}

function valThis(fld, varVar)
{
	varDataType = document.getElementById('varDataType' + varVar).value;
	if ((varDataType == 'float' || varDataType == 'numeric' || varDataType == 'int') && fld.value != '')
	{
		if (!IsNumeric(fld.value))
		{
			alert('<%=getadminRepEditLinksLngStr("DtxtValNumVal")%>');
			fld.value = '';
			fld.focus();
		}
	}
}

function valFrm()
{
	<% If linkType <> "N" and linkType <> "L" and linkType <> "F" and (linkObject = "" or linkObject = "-1") Then
		If linkType = "O" Then linkObjDesc = "Objeto" Else linkObjDesc = "Reporte" %>
	alert('<%=getadminRepEditLinksLngStr("LtxtValSelX")%>'.replace('{0}', '<%=linkObjDesc%>'));
	return false;
	<% ElseIf linkType <> "N" and linkObject <> "" Then
	    If rs.recordcount > 0 Then %>
		if (document.frmLink.varVar.length)
		{
			for (var i = 0;i<document.frmLink.varVar.length;i++)
			{
				if (!valFrmVar(document.frmLink.varVar[i].value, document.frmLink.varNotNull[i].value)) return false;
			}
		}
		else
		{
			if (!valFrmVar(document.frmLink.varVar.value, document.frmLink.varNotNull.value)) return false;
		}
		<% End If %>
	<% End If
	Select Case linkType 
		Case "L" %>
	if (document.frmLink.linkLink.value == '')
	{
		alert('<%=getadminRepEditLinksLngStr("LtxtValLnk")%>');
		return false;
	}
	<% Case "F" %>
	if (document.frmLink.linkObject.selectedIndex <= 0<% If Request("UserType") <> "C" Then %> && document.frmLink.linkObjectPocket.selectedIndex <= 0<% End If %>)
	{
		alert('<%=getadminRepEditLinksLngStr("LtxtValSelForm")%>');
		document.frmLink.linkObject.focus();
		return false;
	}
	<% End Select %>
	document.frmLink.action='adminRepEditLinksSubmit.asp';
	return true;
}

function valFrmVar(varVar, varNotNull)
{
	if (varNotNull == 'Y')
	{
		var rdFld = document.getElementById('rdFld' + varVar);
		var rdVal = document.getElementById('rdVal' + varVar);
		
		if (rdFld.checked)
		{
			if (document.getElementById('valValueF' + varVar).selectedIndex == 0)
			{
				alert('<%=getadminRepEditLinksLngStr("LtxtSelFldVar")%>'.replace('{0}', varVar));
				return false;
			}
		}
		else
		{
			if (document.getElementById('varDataType' + varVar).value != 'datetime')
			{
				if (document.getElementById('valValueV' + varVar).value == '')
				{
					alert('<%=getadminRepEditLinksLngStr("LtxtValVarVal")%>'.replace('{0}', varVar));
					return false;
				}
			}
			else
			{
				if (document.getElementById('colValDat' + varVar).value == '')
				{
					alert('<%=getadminRepEditLinksLngStr("LtxtValVarVal")%>'.replace('{0}', varVar));
					return false;
				}
			}
		}
		return true;
	}
	else return true;
}
</script>

</body>
<% set rs = nothing
conn.close
Function getColTypeVal(ByVal ColType)
If ColType = 129 or ColType = 200 or ColType = 201 or ColType = 130 or ColType = 202 or ColType = 203 Then
	getColTypeVal = "T"
ElseIf ColType = 16 or ColType = 2 or ColType = 3 or ColType = 20 or ColType = 17 or ColType = 18 or _
	ColType = 19 or ColType = 21 or ColType = 4 or ColType = 5 or ColType = 14 or ColType = 131 or ColType = 139 Then
	getColTypeVal = "N"
ElseIf ColType = 7 or ColType = 133 or ColType = 134 or ColType = 135 Then
	getColTypeVal = "D"
Else
	getColTypeVal = "U"
End If
End Function

Sub LoadArrCols
	UserType = Request("UserType")
	set rf = Server.CreateObject("ADODB.RecordSet")
	sql = "select varVar, varDataType from OLKRSvars where rsIndex = " & Request("rsIndex")
	set rf = conn.execute(sql)
	sql = "declare @LanID int set @LanID = 1 "
	do while not rf.eof
		sql = sql & "declare @" & rf("varVar") & " " & rf("varDataType") & " set @" & rf("VarVar") & " = "
		Select Case rf("varDataType")
			Case "nvarchar"
				sql = sql & "'' "
			Case "datetime"
				sql = sql & "'01/01/01' "
			Case "float"
				sql = sql & "0 "
			Case "numeric"
				sql = sql & "0 "
			Case "int"
				sql = sql & "0 "
		End Select
	rf.movenext
	loop
	If repTbl = "OLK" Then
		If UserType = "C" Then
			sql = sql & " declare @CardCode nvarchar(15) set @CardCode = '' "
		ElseIf UserType = "V" Then
			sql = sql & " declare @SlpCode int set @SlpCode = -1 "
		End If
	ElseIf repTbl = "TMRP" Then
		sql = sql & " declare @UserName nvarchar(20) declare @UserID int declare @AlterID nvarchar(100) declare @LanID int "
	End If
	sqlQuery = "select rsQuery, rsTop from OLKRS where rsIndex = " & Request("rsIndex")
	set rf = conn.execute(sqlQuery)
	sqlQuery = rf("rsQuery")
	sql = sql & sqlQuery
	If rf("rsTop") = "Y" Then sql = Replace(sql, "@top", 1)
	sql = QueryFunctions(sql)
	set rf = conn.execute(sql)
	Redim ArrCol(rf.Fields.Count-1)
	Dim ArrColItm(2)
	For i = 0 to rf.Fields.count -1
		myColTypeVal = getColTypeVal(rf(i).Type)
		ArrColItm(0) = rf.Fields(i).Name
		ArrColItm(1) = rf.Fields(i).Name
		ArrColItm(2) = myColTypeVal
		ArrCol(i) = ArrColItm
	next
End Sub %>
</html>
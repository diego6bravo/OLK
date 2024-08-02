<%@ Language=VBScript %>
<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="../chkLogin.asp" -->
<!--#include file="lang/cartEditLine.asp" -->
<!--#include file="../myHTMLEncode.asp"-->
<!--#include file="../authorizationClass.asp"-->
<html <% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>>
<%
Dim myAut
set myAut = New clsAuthorization

Select Case userType
	Case "C"
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKGetCSSPath" & Session("ID")
		cmd.Parameters.Refresh()
		set rs = cmd.execute()
		SelDes = rs(0)
	Case "V"
		SelDes = 0
End Select
 %>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=getcartEditLineLngStr("LttlLineDet")%></title>
<link rel="stylesheet" type="text/css" href="../design/<%=SelDes%>/style/stylePopUp.css">
<link type="text/css" href="../design/0/jquery-ui-1.7.2.custom.css" rel="stylesheet" >	
<script type="text/javascript" src="../jQuery/js/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="../jQuery/js/jquery-ui-1.7.2.custom.min.js"></script>

</head>
<script language="javascript">
var dbID = <%=Session("ID")%>;
var dbName = '<%=Session("olkdb")%>';
var txtValFldMaxChar = "<%=getcartEditLineLngStr("DtxtValFldMaxChar")%>";
var txtValNumVal = "<%=getcartEditLineLngStr("DtxtValNumVal")%>";
var txtValNumMinVal = "<%=getcartEditLineLngStr("DtxtValNumMinVal")%>";
var txtValNumMaxVal = "<%=getcartEditLineLngStr("DtxtValNumMaxVal")%>";
var txtValNumValWhole = "<%=getcartEditLineLngStr("DtxtValNumValWhole")%>";
</script>
<script language="javascript" src="cartEditLine.js"></script>
<script type="text/javascript" src="../scr/calendar.js"></script>
<script type="text/javascript" src="../scr/lang/calendar-<%=Left(Session("myLng"), 2)%>.js"></script>
<script type="text/javascript" src="../scr/calendar-setup.js"></script>
<link rel="stylesheet" type="text/css" media="all" href="../design/0/style/style_cal.css" title="winter">
<script language="javascript" src="../general.js"></script>
<script language="javascript" src="../generalData.js.asp?dbID=<%=Session("ID")%>&LastUpdate=<%=myApp.LastUpdate%>"></script>
<body topmargin="0" leftmargin="0" onfocus="javascript:chkWin();">
<%
set rd = Server.CreateObject("ADODB.RecordSet")
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKCheckRestoreUDF" & Session("ID")
cmd.Parameters.Refresh()
cmd("@SysID") = "INV1"
cmd("@ObsID") = "DOC1"
set rs = cmd.execute()
If rs(0) = "Y" Then %>
<script language="javascript">
opener.location.href='../configErr.asp?errCmd=DocLines';
window.close();
</script>
<% Else

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetCartLineData" & Session("ID")
cmd.Parameters.Refresh()
cmd("@LogNum") = Session("RetVal")
cmd("@LineNum") = Request("LineNum")
cmd("@LanID") = Session("LanID")
set rs = cmd.execute()

If Request("cmd") = "u" then
	If myApp.SDKLineMemo then LineMemo = RS("LineMemo")
	WhsCode = RS("WhsCode")
Else
	If myApp.SDKLineMemo then LineMemo = Request("LineMemo")
	WhsCode = Request("WhsCode")
End If

If Request("cmd") <> "e" Then
	VatGroup = rs("VatGroup")
	TaxCode = rs("TaxCode")
Else
	VatGroup = Request("VatGroup")
	TaxCode = Request("TaxCode")
End If

TreeType = rs("TreeType")

ObjCode = rs("ObjCode")
			
set rw = Server.CreateObject("ADODB.recordset")
%>
<form method="POST" action="cartEditLineUpdate.asp" name="form1">
<div align="left">
	<table border="0" cellpadding="0" width="100%">
	<% If userType = "V" Then %>
		<tr class="GeneralTlt">
			<td colspan="2">
			<%=getcartEditLineLngStr("LttlLineDet")%></td>
		</tr>
      <%  If (TreeType <> "S" and TreeType <> "C" or TreeType = "S" and myApp.TreePricOn) Then
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKGetWarehouses" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@LanID") = Session("LanID")
      		If myAut.HasAuthorization(99) Then
			set rw = cmd.execute() %>
		<tr class="GeneralTbl">
			<td class="GeneralTblBold2" width="25%">
			<%=getcartEditLineLngStr("DtxtWarehouse")%></td>
			<td><select size="1" name="whscode" style="width: 100%">
		    <% do while not rw.eof %>
			<option value="<%=myHTMLEncode(RW("WhsCode"))%>" <% If Rw("WhsCode") = WhsCode Then %>selected<%end if %>><%=myHTMLEncode(RW("WhsName"))%></option>
			<% rw.movenext
			loop %>
			</select></td>
		</tr>
		<% Else 
			cmd("@Filter") = WhsCode
			set rw = cmd.execute() %>
		<tr class="GeneralTbl">
			<td class="GeneralTblBold2" width="25%">
			<%=getcartEditLineLngStr("DtxtWarehouse")%></td>
			<td><input type="hidden" name="whscode" value="<%=WhsCode%>"><%=myHTMLEncode(RW("WhsName"))%></td>
		</tr>
		<% End If %>
      <% Select Case myApp.LawsSet 
	      	Case "PA", "AT", "AU", "BE", "CH", "CZ", "DE", "DK", "ES", "FI", "FR", "HU", "IT", "NL", "NO", "PL", "PT", "RU", "SE", "SK", "GB", "ZA" %>
		<tr class="GeneralTbl">
			<td class="GeneralTblBold2" width="25%">
			<%=getcartEditLineLngStr("LtxtVatGrp")%></td>
			<td>
			<% 
			Cat = "O"
			If ObjCode = 22 Then Cat = "I"
			
			If myAut.HasAuthorization(175) Then
			
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKGetVatGroup" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@LanID") = Session("LanID")
			cmd("@Category") = Cat
			set rw = cmd.execute() %>
	        <select size="1" name="VATGroup" style="width: 100%">
		    <% do while not rw.eof %>
			<option value="<%=myHTMLEncode(RW(0))%>" <% If Rw(0) = VATGroup Then %>selected<%end if %>><%=myHTMLEncode(RW(1))%></option>
			<% rw.movenext
			loop %>
	        </select><% Else
	        
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKGetVatGroup" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@LanID") = Session("LanID")
			cmd("@Category") = Cat
			cmd("@Filter") = VATGroup
			set rw = cmd.execute()
	        Response.Write rw("Name") %><input type="hidden" name="VATGroup" value="<%=myHTMLEncode(VATGroup)%>"><% End If %></td>
		</tr>
		<% Case "MX", "CL", "CR", "GT", "US", "CA" %>
		<tr class="GeneralTbl">
			<td class="GeneralTblBold2" width="25%">
			<%=getcartEditLineLngStr("LtxtTaxCode")%></td>
			<td><% If myAut.HasAuthorization(175) Then %>
	        <select size="1" name="TaxCode" style="width: 100%">
	        <option></option>
			<% 
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKGetTaxCodes" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@LanID") = Session("LanID")
			If ObjCode = 22 Then cmd("@Category") = "S"
			set rw = cmd.execute()
			do while not rw.eof %>
			<option <% If TaxCode = rw("Code") Then %>selected<% End If %> value="<%=myHTMLEncode(rw("Code"))%>"><%=myHTMLEncode(rw("Code"))%> 
			- <%=myHTMLEncode(rw("Name"))%></option>
			<% rw.movenext
			loop %>
			</select><% Else
				set cmd = Server.CreateObject("ADODB.Command")
				cmd.ActiveConnection = connCommon
				cmd.CommandType = &H0004
				cmd.CommandText = "DBOLKGetTaxCodes" & Session("ID")
				cmd.Parameters.Refresh()
				cmd("@LanID") = Session("LanID")
				If ObjCode = 22 Then cmd("@Category") = "S"
				cmd("@Filter") = TaxCode
		        set rw = cmd.execute()
		        If Not rw.Eof Then Response.Write TaxCode %><input type="hidden" name="TaxCode" value="<%=myHTMLEncode(TaxCode)%>"><% 
	       	 End If
	      %></td>
		</tr>
      <% End Select
      Else
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKGetWarehouses" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@LanID") = Session("LanID")
			cmd("@Filter") = WhsCode
			set rw = cmd.execute() %>
		<tr class="GeneralTbl">
			<td class="GeneralTblBold2" width="25%">
			<%=getcartEditLineLngStr("DtxtWarehouse")%></td>
			<td><%=rw("WhsName")%><input type="hidden" name="whscode" value="<%=myHTMLEncode(WhsCode)%>"></td>
		</tr>
      <% 
      Select Case myApp.LawsSet 
	      	Case "PA", "AT", "AU", "BE", "CH", "CZ", "DE", "DK", "ES", "FI", "FR", "HU", "IT", "NL", "NO", "PL", "PT", "RU", "SE", "SK", "GB", "ZA" %>
		<tr class="GeneralTbl">
			<td class="GeneralTblBold2" width="25%">
			<%=getcartEditLineLngStr("LtxtVatGrp")%></td>
			<td>
	        <% 
			Cat = "O"
			If ObjCode = 22 Then Cat = "I"
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKGetVatGroup" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@LanID") = Session("LanID")
			cmd("@Category") = Cat
			cmd("@Filter") = VATGroup
			set rw = cmd.execute() %>
	        <%=rw("Name")%>
	        <input type="hidden" name="VATGroup" value="<%=myHTMLEncode(VATGroup)%>"></td>
		</tr>
		<% Case "MX", "CL", "CR", "GT", "US", "CA" %>
		<tr class="GeneralTbl">
			<td class="GeneralTblBold2" width="25%">
			<%=getcartEditLineLngStr("LtxtTaxCode")%></td>
			<td>
			<% 
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKGetTaxCodes" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@LanID") = Session("LanID")
			If ObjCode = 22 Then cmd("@Category") = "S"
			cmd("@Filter") = TaxCode
			set rw = cmd.execute()
			If Not rw.Eof Then Response.Write rw("Name")%>
     		<input type="hidden" name="TaxCode" value="<%=myHTMLEncode(TaxCode)%>"></td>
		</tr>
      <% End Select %>
      <% End If %>
		<tr class="GeneralTbl">
			<td class="GeneralTblBold2" width="25%"><%=getcartEditLineLngStr("DtxtNote")%></td>
			<td>
			<select <% If Not myApp.SDKLineMemo Then %>disabled<% end if %> size="1" name="NoteVar" style="width: 100%" onchange="updateNote()">
			<option><%=getcartEditLineLngStr("LtxtPredNotes")%></option>
			<% set rw = conn.execute("select NoteName, Note from OLKNotePerLine")
			do while not rw.eof %>
			<option value="<%=myHTMLEncode(RW("Note"))%>"><%=myHTMLEncode(RW("NoteName"))%></option>
			<% rw.movenext
			loop %>
			</select></td>
		</tr>
		<tr class="GeneralTbl">
			<td  class="GeneralTblBold2" width="25%"><%=getcartEditLineLngStr("DtxtNote")%></td>
			<td><textarea <% If Not myApp.SDKLineMemo Then %>disabled<% end if %> rows="5" name="LineMemo" cols="47" onkeydown="return chkMax(event, this, <% If myApp.SVer < 8 Then %>254<% Else %>64000<% End If %>);" style="width: 100%; "><%=myHTMLEncode(LineMemo)%><% If Not myApp.SDKLineMemo Then %><%=getcartEditLineLngStr("LtxtDisNotes")%><% end if %></textarea></td>
		</tr>
		<% 
		End If
		set rSdk = Server.CreateObject("ADODB.RecordSet")
		set rg = Server.CreateObject("ADODB.RecordSet")
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKGetUDFGroups" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@UserType") = userType
		cmd("@TableID") = "INV1"
		cmd("@OP") = "O"
		rg.open cmd, , 3, 1

		If not rg.eof Then

			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKGetUDFWriteCols" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@LanID") = Session("LanID")
			cmd("@UserType") = userType
			cmd("@TableID") = "INV1"
			cmd("@OP") = "O"
			rSdk.open cmd, , 3, 1
			set rd = Server.CreateObject("ADODB.RecordSet")
			
			do while not rg.eof %>
		<tr class="GeneralTlt">
			<td colspan="2">
			<% Select Case CInt(rg("GroupID"))
				Case -1 %><%=getcartEditLineLngStr("DtxtUDF")%><%
				Case Else
					Response.Write rg("GroupName")
				End Select %></td>
		</tr><%	
			rSdk.Filter = "GroupID = " & rg("GroupID")	
			do while not rSdk.eof
				InsertID = rSdk("InsertID")
				If Request("cmd") <> "e" Then
					FldVal = rs(InsertID)
				Else
					If Request("U_" & rSdk("AliasID")) <> "" Then
						FldVal = Request("U_" & rSdk("AliasID"))
					Else
						FldVal = rSdk("NullVal")
					End If
				End If %>
				<tr class="generalTbl">
			            <td width="25%" class="GeneralTblBold2">
			              <table border="0" cellpadding="0" cellspacing="0" width="100%">
			                <tr class="GeneralTblBold2">
			            	  <td>
			            	    <b><font size="1" face="Verdana"><%=rSdk("Descr")%><% If rSdk("NullField") = "Y" Then %><font color="red">*</font><% End If %></font></b></td>
			            	    <% If (rSdk("Query") = "Y" or rSdk("TypeID") = "D") and IsNull(rSdk("RTable")) Then %>
			            	    <td width="16">
			            	    	<img border="0" src="../images/<% If rSdk("TypeID") <> "D" Then %>flechaselec2<% Else %>cal<% End If %>.gif" id="btn<%=rSdk("AliasID")%>" <% If rSdk("Query") = "Y" Then %>onclick="datePicker('../SmallQuery.asp?sType=DocLine&LineNum=<%=Request("LineNum")%>&FieldID=<%=rSdk("FieldID")%>&pop=Y<% If rSdk("TypeID") = "A" Then %>&MaxSize=<%=rSdk("SizeID")%><% End If %>',400,250,'yes', 'yes', document.form1.U_<%=rSdk("AliasID")%>)"<% End If %>>
			            	    </td>
			            	    <% End If %>
			            	</tr>
			              </table>
			            </td>
			            <td dir="ltr" bgcolor="#EAF5FF"><% If rSdk("DropDown") = "Y" or not IsNull(rSdk("RTable")) then 
				            	set rd = Server.CreateObject("ADODB.RecordSet")
				            	If rSdk("DropDown") = "Y" Then
									cmd.CommandText = "DBOLKGetUDFValues" & Session("ID")
									cmd.Parameters.Refresh()
									cmd("@LanID") = Session("LanID")
									cmd("@TableID") = "INV1"
									cmd("@FieldID") = rSdk("FieldID")
									rd.open cmd, , 3, 1
								  Else
								  	sql = "select Code, Name from [@" & rSdk("RTable") & "] order by 2"
								  	rd.open sql, conn, 3, 1
								  End If
							 %><select size="1" name="U_<%=rSdk("AliasID")%>" class="input" style="width: 99%">
								<option></option>
								<% do while not rd.eof %>
								<option <% If Not IsNull(FldVal) Then If CStr(FldVal) = CStr(rd(0)) Then Response.Write "Selected" %> value="<%=rd(0)%>" <% If rSdk("Dflt")= rd(0) Then %>selected<% End If %>><%=myHTMLEncode(rd(1))%></option>
								<% rd.movenext
								loop
								rd.close %>
							</select>
					<% ElseIf rSdk("TypeID") = "M" and Trim(rSdk("EditType")) = "" or rSdk("TypeID") = "A" and rSdk("EditType") = "?" Then %>
						<% If rSdk("Query") = "Y" or rSdk("TypeID") = "D" Then %>
						<table width="100%" cellspacing="0" cellpadding="0">
						  <tr>
						    <td>
						<% End If %>
						<textarea <% If rSdk("TypeID") = "D" or rSdk("Query") = "Y" Then %>readonly<% End If %> type="text" name="U_<%=rSdk("AliasID")%>" class="input" onchange="chkThis(this, '<%=rSdk("TypeID")%>', '<%=rSdk("EditType")%>', <%=rSdk("SizeID")%>)" <% If rSdk("Query") = "Y" Then %>onclick="datePicker('../SmallQuery.asp?sType=DocLine&LineNum=<%=Request("LineNum")%>&FieldID=<%=rSdk("FieldID")%>&pop=Y<% If rSdk("TypeID") = "A" Then %>&MaxSize=<%=rSdk("SizeID")%><% End If %>',500,300,'yes', 'yes', this)"<% End If %> rows="3" onfocus="this.select()" style="width: 100%" cols="1"><% If Not IsNull(FldVal) Then %><%=myHTMLEncode(FldVal)%><% Else %><% If Not IsNull(rSdk("Dflt")) Then %><%=rSdk("Dflt")%><% End If %><% End If %></textarea>
						<% If rSdk("Query") = "Y" or rSdk("TypeID") = "D" Then %>
							</td>
							<td width="16">
								<img border="0" src="../images/<%=Session("rtl")%>remove.gif" width="16" height="16" onclick="document.form1.U_<%=rSdk("AliasID")%>.value = ''" style="cursor: hand">
							</td>
						  </tr>
						</table>
						<% End If %>
					<% ElseIf rSdk("TypeID") = "A" and rSdk("EditType") = "I" Then %>
						<table cellpadding="0" cellspacing="2" border="0">
							<tr>
								<td><img src="../pic.aspx?filename=<% If IsNull(FldVal) Then %>n_a.gif<% Else %><%=FldVal%><% End If %>&MaxSize=180&dbName=<%=Session("olkdb")%>" id="imgU_<%=rSdk("AliasID")%>" border="1">
								<input type="hidden" name="U_<%=rSdk("AliasID")%>" value="<%=Trim(FldVal)%>"></td>
								<td width="16" valign="bottom"><img border="0" src="../images/<%=Session("rtl")%>remove.gif" width="16" height="16" onclick="javascript:document.form1.U_<%=rSdk("AliasID")%>.value = '';document.form1.imgU_<%=rSdk("AliasID")%>.src='../pic.aspx?filename=n_a.gif&MaxSize=180&dbName=<%=Session("olkdb")%>';" style="cursor: hand"></td>
							</tr>
							<tr>
								<td colspan="2" height="22">
								<p align="center">
								<input type="button" value="<%=getcartEditLineLngStr("DtxtAddImg")%>" name="B1" onclick="javascript:getImg(document.form1.U_<%=rSdk("AliasID")%>, document.form1.imgU_<%=rSdk("AliasID")%>,180);"></td>
							</tr>
						</table>
						<% Else
						If Not IsNull(FldVal) Then 
							If rSdk("TypeID") = "B" Then
				        	Select Case rSdk("EditType")
								Case "R"
									FldVal = FormatNumber(CDbl(FldVal),myApp.RateDec)
								Case "S"
									FldVal = FormatNumber(CDbl(FldVal),myApp.SumDec)
								Case "P"
									FldVal = FormatNumber(CDbl(FldVal),myApp.PriceDec)
								Case "Q"
									FldVal = FormatNumber(CDbl(FldVal),myApp.QtyDec)
								Case "%"
									FldVal = FormatNumber(CDbl(FldVal),myApp.PercentDec)
								Case "M"
									FldVal = FormatNumber(CDbl(FldVal),myApp.MeasureDec)
				        	End Select
				        	End If
						Else
							FldVal = ""
						End If %>
							<% If rSdk("Query") = "Y" or rSdk("TypeID") = "D" Then %>
							<table width="100%" cellspacing="0" cellpadding="0">
							  <tr>
							    <td>
							<% End If %>
							<% 
							If rSdk("TypeID") = "D" or rSdk("Query") = "Y" Then readOnly = True Else readOnly = False
							If rSdk("TypeID") = "D" Then FldVal = FormatDate(FldVal, False)
							If rSdk("TypeID") = "A" Then fldSize = 43 Else fldSize = 12
							If rSdk("TypeID") = "B" or rSdk("TypeID") = "A" Then
								If rSdk("TypeID") = "B" Then MaxSize = 21 Else MaxSize = rSdk("SizeID")
								isMaxSize = True
							Else
								isMaxSize = False
							End If %>
							<input <% If readOnly Then %>readonly<% End If %> type="text" name="U_<%=rSdk("AliasID")%>" id="U_<%=rSdk("AliasID")%>" size="<%=fldSize%>" class="input" <% If rSdk("TypeID") = "N" or rSdk("TypeID") = "B" Then %>onkeydown="return valKeyNum<% If rSdk("TypeID") = "V" Then %>Dec<% End If %>(event);"<% End If %> <% If rSdk("TypeID") = "D" Then %>onclick="btn<%=rSdk("AliasID")%>.click();"<% End If %> onchange="chkThis(this, '<%=rSdk("TypeID")%>', '<%=rSdk("EditType")%>', <%=rSdk("SizeID")%>)" <% If rSdk("TypeID") = "D" Then %>onclick="btn<%=rSdk("AliasID")%>.click()"<% End If %> <% If rSdk("Query") = "Y" Then %>onclick="datePicker('../SmallQuery.asp?sType=DocLine&LineNum=<%=Request("LineNum")%>&FieldID=<%=rSdk("FieldID")%>&pop=Y<% If rSdk("TypeID") = "A" Then %>&MaxSize=<%=rSdk("SizeID")%><% End If %>',500,300,'yes', 'yes', this)"<% End If %> value="<% If Not IsNull(FldVal) Then %><%=FldVal%><% Else %><% If Not IsNull(rSdk("Dflt")) Then %><%=rSdk("Dflt")%><% End If %><% End If %>" <% If rSdk("TypeID") <> "D" Then %>onfocus="this.select()"<% End If %> style="width: 100%" <% If isMaxSize Then %> onkeydown="return chkMax(event, this, <%=MaxSize%>);"<% End if %>>
							<% If rSdk("Query") = "Y" or rSdk("TypeID") = "D" Then %>
								</td>
								<td width="16">
									<img border="0" src="../images/<%=Session("rtl")%>remove.gif" width="16" height="16" onclick="document.form1.U_<%=rSdk("AliasID")%>.value = ''">
								</td>
							  </tr>
							</table>
							<% End If %>
						<% End If %>
			            </td>
			          </tr>
		<%	rSdk.movenext
			loop
		rg.movenext
		loop
		End If  %>
		<tr class="GeneralTbl">
			<td colspan="2">
			<table cellpadding="0" cellspacing="0" border="0" width="100%">
				<tr>
					<td><input type="submit" value="<%=getcartEditLineLngStr("DtxtSave")%>" name="B1"></td>
					<td align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><input type="button" value="<%=getcartEditLineLngStr("DtxtCancel")%>" name="B2" onclick="window.close();"></td>
				</tr>
			</table>
			</td>
		</tr>
	</table>
</div>
<input type="hidden" name="LineNum" value="<%=Request("LineNum")%>">
<input type="hidden" name="redir" value="<%=Request("redir")%>">
<input type="hidden" name="AddPath" value="../">
<input type="hidden" name="pop" value="Y">
</form>
<script language="javascript">
<% 
If rg.recordcount > 0 Then
	rSdk.Filter = "TypeID = 'D'"
	If rSdk.recordcount > 0 Then rSdk.movefirst
	do while not rSdk.eof %>
	    Calendar.setup({
	        inputField     :    "U_<%=rSdk("AliasID")%>",     // id of the input field
	        ifFormat       :    "<%=GetCalendarFormatString%>",      // format of the input field
	        button         :    "btn<%=rSdk("AliasID")%>",  // trigger for the calendar (button ID)
	        align          :    "Bl",           // alignment (defaults to "Bl")
	        singleClick    :    true
	    });
	<% rSdk.movenext
	loop %>
<% End If %>
<% If Request("cmd") = "e" Then %>
alert("<%=getcartEditLineLngStr("LtxtValItmQty")%>")
<% end if %>
</script>
<% End If %>
</body>
<% conn.close %>
</html>
<!--#include file="clientInc.asp"-->
<% 
ObjID = CInt(Request("adObjID"))
adSearchID = CInt(Request("ID"))

Select Case userType
Case "C" %><!--#include file="clientTop.asp"-->
<% Case "V" %><!--#include file="agentTop.asp"-->
<% 
If (ObjID = 2 and not myAut.HasBPAccess or ObjID = 4 and not myAut.HasAuthorization(1)) or (ObjID <> 2 and ObjID <> 4) Then Response.Redirect "unauthorized.asp"
End Select %>
<% addLngPathStr = "" %>
<!--#include file="lang/adCustomSearch.asp" -->
<script language="javascript" src="js_up_down.js"></script>
<script type="text/javascript" src="scr/calendar.js"></script>
<script type="text/javascript" src="scr/lang/calendar-<%=Left(Session("myLng"), 2)%>.js"></script>
<script type="text/javascript" src="scr/calendar-setup.js"></script>
<% 

If userType = "V" Then
	QryGroup = myApp.CarArt
End If

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetAdCustomSearch" & Session("ID")
cmd.Parameters.Refresh()
cmd("@LanID") = Session("LanID")
cmd("@ObjID") = ObjID
cmd("@ID") = adSearchID
set rs = cmd.execute()

CatType = rs("CatType")
If CatType = "S" Then CatType = "T"
Order1 = rs("Order1")
Order2 = rs("Order2")
	
If Order1 = "" or IsNull(Order1) Then
	Select Case myApp.GetDefCatOrdr
		Case "C"
			Order1 = "OITM.ItemCode"
		Case "N"
			Order1 = "ItemName"
	End Select
End If

If Order2 = "" or IsNull(Order2) Then
	Order2 = "A"
End If

set rd = server.createobject("ADODB.RecordSet")

strFromTo = " " & getadCustomSearchLngStr("DtxtFrom") & " - " & getadCustomSearchLngStr("DtxtTo")
%>
<form method="POST" action="<% If ObjID = 2 Then %>clientsSearch.asp<% Else %>search.asp<% End If %>" name="frmVars" onsubmit="javascript:return doLoadDesc();">
<div align="center">
	<table border="0" cellpadding="0" width="100%" style="font-family: Verdana; font-size: 10px">
		<% If tblCustTtl = "" Then %>
		<tr>
			<td class="TablasTituloSec" id="tdMyTtl">&nbsp;<%=rs("Name")%></td>
		</tr>
		<% Else %>
		<%=Replace(tblCustTtl, "{txtTitle}", rs("Name"))%>
		<% End If %>
		<tr class="TablasNoticias">
			<td>
			&nbsp;</td>
		</tr>
		<tr>
			<td>
			<table border="0" cellpadding="0" width="100%">
			<%
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKGetAdCustomSearchDetils" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@LanID") = Session("LanID")
			cmd("@ObjID") = ObjID
			cmd("@ID") = adSearchID
			If Session("UserName") <> "" Then cmd("@CardCode") = Session("UserName")
			rs.close
			rs.open cmd, , 3, 1
			set rQVal = Server.CreateObject("ADODB.RecordSet")
			set rBase = Server.CreateObject("ADODB.RecordSet")
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKGetAdCustomSearchBase" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@LanID") = Session("LanID")
			cmd("@ObjID") = ObjID
			cmd("@ID") = adSearchID
			rBase.open cmd, , 3, 1
			hdCatType = True
			hdOrder = True
			do while not rs.eof 
				enableControl = True
				If rs("NotNull") = "Y" Then
					If notNullVars <> "" Then notNullVars = notNullVars & ", "
					notNullVars = notNullVars & "var" & rs("varID") & "~" & rs("Name") & "~" & rs("Type")
				End If
				Select Case rs("Type")
					Case "DD", "CL", "L"
					   If rs("DefVars") = "F" Then
					   		sql = "select T0.Value, IsNull(T1.alterDescription, T0.Description) Description " & _
					   				"from OLKCustomSearchVarsVals T0 " & _
					   				"left outer join OLKCustomSearchVarsValsAlterNames T1 on T1.ObjectCode = T0.ObjectCode and T1.ID = T0.ID and T1.varID = T0.varID and T1.ValID = T0.ValID and T1.LanID = " & Session("LanID") & " " & _
					   				"where T0.ObjectCode = " & ObjID & " and T0.ID = " & adSearchID & " and T0.varID = " & rs("varID")
					   Else
							sql = getCustomSearchValsCSQL(true, rs("Query"))
					   End If
					Case "Q"
						If rs("DefVars") = "Q" Then
							sql = getCustomSearchValsCSQL(false, "")
						End If
				End Select
				If Request.Form("isSubmit") <> "R" Then
					If rs("Type") <> "DP" and rs("DefValBy") = "V" Then
						defValue = rs("DefValValue")
					ElseIf rs("Type") = "DP" and rs("DefValBy") = "V" Then
						defValue = FormatDate(rs("DefValDate"), False)
					ElseIf rs("DefValBy") = "Q" Then
						sqlVal = getCustomSearchValsCSQL(true, rs("DefValValue"))
						set rQVal = conn.execute(sqlVal)
						If not rQVal.eof then
							If rs("Type") = "DP" Then
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
					defValue = Request.Form("var" & rs("varID"))
				End If  %>
				<tr>
					<td width="160" class="TablasNoticias">
					<table cellpadding="0" cellspacing="0" border="0" width="100%">
					<tr>
						<td class="TablasNoticias" style="vertical-align: top; padding-top: 2px;" valign="top"><% Select Case rs("Type")
						Case "S"
							Select Case rs("Variable")
								Case "Search" 
									Response.Write getadCustomSearchLngStr("DtxtSearch")
								Case "ItmsGrpCod", "ItmsGrpRange"
									Response.Write txtAlterGrp
									If rs("Variable") = "ItmsGrpRange" Then Response.Write strFromTo
								Case "FirmCode", "FirmRange"
									Response.Write txtAlterFrm
									If rs("Variable") = "FirmRange" Then Response.Write strFromTo
								Case "Order" 
									Response.Write getadCustomSearchLngStr("DtxtOrder")
								Case "PriceRange" 
									Response.Write getadCustomSearchLngStr("DtxtPrice") & strFromTo
								Case "Inventory"
									Response.Write getadCustomSearchLngStr("LtxtInvMoreThen") 
								Case "InvRange"
									Response.Write getadCustomSearchLngStr("LtxtInventory") & strFromTo
								Case "CatType"
									Response.Write getadCustomSearchLngStr("LtxtCatalogType")
								Case "ItemRange"
									Response.Write getadCustomSearchLngStr("DtxtItem") & strFromTo
								Case "CardType"
									Response.Write getadCustomSearchLngStr("DtxtType")
								Case "BPRange"
									Response.Write getadCustomSearchLngStr("DtxtBPCode") & strFromTo
								Case "BPGrpRange"
									Response.Write getadCustomSearchLngStr("DtxtGroup") & strFromTo
								Case "BPCntRange"
									Response.Write getadCustomSearchLngStr("DtxtCountry") & strFromTo
								Case "ItmProp", "BPProp"
									Response.Write getadCustomSearchLngStr("DtxtProp")									
								Case Else
									Response.Write "&nbsp;"
							End Select
						Case Else
							Response.Write rs("Name")
						End Select %></td>
						<% Select Case rs("Type") 
							Case "Q" %><td width="15"><img border="0" src="design/<%=SelDes%>/images/<%=Session("rtl")%>flecha_selec.gif" <% If enableControl Then %>onclick="javascript:Start(document.frmVars.var<%=rs("varID")%>,'SmallQuery.asp?source=customsearch&ObjID=<%=ObjID%>&ID=<%=adSearchID%>&varID=<%=rs("varID")%>&s=<%=rs("DefVars")%>',550,250,'Yes','Yes');"<% End If %>></td>
						<% Case "DP" %>
							<td width="16"><img border="0" src="images/cal.gif" width="16" height="16" id="btn<%=rs("varID")%>"></td>
						<% End Select %>
					</tr>
					</table>
					</td>
					<td class="TblGeneral"><% Select Case rs("Type")
				   Case "DD" %><select <% If Not enableControl Then %>disabled<% End If %> name="var<%=rs("varID")%>" size="1" <% If rs("IsBase") = "Y" Then %> onchange="reload(<%=rs("TargetID")%>);"<% End If %> style="width: 300px; ">
				   <% If enableControl Then %>
				   <option></option><%
				   set rd = conn.execute(sql)
				   do while not rd.eof
				   %><option <% If defValue = CStr(rd(0)) Then %>selected<% End If %> value="<%=rd(0)%>"><%=myHTMLEncode(rd(1))%></option>
				   <% rd.movenext
				   loop
				   Else %>
				   <option value=""><%=getadCustomSearchLngStr("LtxtSelectEnter")%> "<%=selectRepValsCName%>"</option>
				   <% End If %>
				   </select>
				<% Case "T" %><input type="text" name="var<%=rs("varID")%>" id="var<%=rs("varID")%>" size="20" onchange="chkNum(this, '<%=rs("DataType")%>');<% If rs("IsBase") = "Y" Then %>reload(<%=rs("TargetID")%>);<% End If %>" value="<%=defValue%>" style="width: 300px; ">
				<% Case "L" %><select <% If Not enableControl Then %>disabled<% End If %> name="var<%=rs("varID")%>" size="5" <% If rs("IsBase") = "Y" Then %> onchange="reload(<%=rs("TargetID")%>);"<% End If %> style="width: 300px; "><%
					If enableControl Then
					set rd = conn.execute(sql)
				   do while not rd.eof
				   %><option <% If defValue = CStr(rd(0)) Then %>selected<% End If %> value="<%=rd(0)%>"><%=myHTMLEncode(rd(1))%></option>
				   <% rd.movenext
				   loop
				   Else %>
				   <option value=""><%=getadCustomSearchLngStr("LtxtSelectEnter")%> "<%=selectRepValsCName%>"</option>
				   <% End If %>
				</select>
				<% Case "Q" %><input readonly type="text" name="var<%=rs("varID")%>" size="16" <% If enableControl Then %>onclick="javascript:Start(this, 'SmallQuery.asp?source=customsearch&ObjID=<%=ObjID%>&ID=<%=adSearchID%>&varID=<%=rs("varID")%>&s=<%=rs("DefVars")%>',550,250,'Yes','Yes');" value="<%=defValue%>" <% Else %> disabled value="<%=getadCustomSearchLngStr("LtxtSelectEnter")%> &quot;<%=selectRepValsCName%>&quot;" <% End If %> <% If rs("IsBase") = "Y" Then %> onchange="reload(<%=rs("TargetID")%>);"<% End If %> style="width: 300px; ">
				<% Case "DP" %><table border="0" cellpadding="0" cellspacing="0"><tr>
				<td>
				<input name="var<%=rs("varID")%>" id="var<%=rs("varID")%>" size="12" readonly value="<%=defValue%>" <% If rs("Type") = "DP" Then %>onclick="btn<%=rs("varID")%>.click();"<% End If %> <% If rs("IsBase") = "Y" Then %> onchange="reload(<%=rs("TargetID")%>);"<% End If %>></td></tr></table>
				<% Case "CL" 
				   If enableControl Then %>
					<ilayer name="scroll<%=rs("varID")%>1" width=100% height=120 clip="0,0,170,150">
					<layer name="scroll<%=rs("varID")%>2" width=100% height=120 bgColor="white">
					<div id="scroll<%=rs("varID")%>3" style="width:100%;height:120px;overflow:auto">
				   	<% 
				   	set rd = conn.execute(sql)
				   	doChk = False
				   	If Request.Form("var" & rs("VarID")) <> "" Then
				   		set rdChk = Server.CreateObject("ADODB.RecordSet")
				   		sql = "select * from OLKCommon.dbo.OLKSplit(N'" & saveHTMLDecode(Request.Form("var" & rs("VarID")), False) & "', ', ') "
				   		rdChk.open sql, conn, 3, 1
				   		doChk = True
				   	End If
				   	i = 0
				   	do while not rd.eof
				   	chk = False
				   	If doChk Then
				   		rdChk.Filter = "Value = '''" & saveHTMLDecode(rd(0), False) & "'''"
				   		chk = rdChk.recordcount > 0
				   	End If
				   	 %>
					<input type="checkbox" name="var<%=rs("varID")%>" <% If chk Then %>checked<% End If %> value="'<%=Replace(myHTMLEncode(rd(0)), "'", "''")%>'" id="var<%=rs("varID")%><%=i%>"  style="border-style:solid; border-width:0; background:background-image" <% If rs("IsBase") = "Y" Then %> onclick="reload(<%=rs("TargetID")%>);"<% End If %>><label id="txt<%=rs("varID")%><%=i%>" for="var<%=rs("varID")%><%=i%>"><%=myHTMLEncode(rd(1))%></label><br>
					<% i = i + 1
					rd.movenext
					loop
					Else %>
				   <label><%=getadCustomSearchLngStr("LtxtSelectEnter")%></label> "<%=selectName%>"
					</div>
					</layer>
					</ilayer>
				   <% End If %>
				   <% Case "S"
				   		Select Case rs("Variable")
				   			Case "Search" %><table cellpadding="0" cellspacing="0" border="0">
				   				<tr>
				   					<td><input type="text" name="string" size="29" value="" style="width: 270px;"></td>
				   				</tr>
								<% If myApp.SearchExactA Then %>
								<tr>
									<td>
									<p align="center">
									<font face="Verdana" size="1">
									<input type="radio" value="E" name="rdSearchAs" class="noborder" id="rdSearchAsE" <% If Request("rdSearchAs") = "" and myApp.SearchMethodA = "E" or Request("rdSearchAs") = "E" Then %>checked<% End If %>><label for="rdSearchAsE"><%=getadCustomSearchLngStr("DtxtExact")%></label>
									<input type="radio" name="rdSearchAs" class="noborder" id="rdSearchAsS" value="S" <% If Request("rdSearchAs") = "" and myApp.SearchMethodA = "L" or Request("rdSearchAs") = "S" Then %>checked<% End If %>><label for="rdSearchAsS"><%=getadCustomSearchLngStr("DtxtLike")%></label></font>
									</td>
								</tr>
								<% Else %>
								<input type="hidden" name="rdSearchAs" value="S">
								<% End If %>
				   			</table>
				   		<% Case "ItmsGrpCod"
				   		
							innerAddStr = ""
							If myApp.GetMinInvBy = "W" Then
								innerAddStr = "inner join OITW T____1 on T____1.ItemCode = T____0.ItemCode and T____1.WhsCode = OLKCommon.dbo.DBOLKGetWhsCode" & Session("ID") & "(" & Session("branch") & ", " & Session("vendid") & ", T____0.ItemCode) "
								fldTbl = "T____1"
							Else
								fldTbl = "T____0"
							End If
							
							sql = "select X____0.ItmsGrpCod, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITB', 'ItmsGrpNam', X____0.ItmsGrpCod, X____0.ItmsGrpNam) ItmsGrpNam  " & _
							"from OITB X____0 "
						
							If myApp.ApplyInvFiltersBy = "N" Then	
								sql = sql & "where X____0.ItmsGrpCod in "
							ElseIf myApp.ApplyInvFiltersBy = "I" Then
								sql = sql & " inner join "
							End If
							
							sql = sql & "(select ItmsGrpCod from oitm T____0 " & innerAddStr & " where SellItem = 'Y' and " & fldTbl & ".OnHand >= " & myApp.GetMinInv & " "
							
							If QryGroup <> "-1" Then sql = sql & " and QryGroup" & QryGroup & " = 'N' "
							
							If Not IsNull(CatalogFilter) and CatalogFilter <> "" Then
								sql = sql & " and T____0.ItemCode not in (" & CatalogFilter & ") "
							End If
							
							If myApp.GetApplyGenFilter Then
								sql = sql & " and T____0.ItemCode not in (" & myApp.GetGenFilter & ") "
							End If
							
							If Not IsNull(searchTreeFilter) and searchTreeFilter <> "" Then
								sql = sql & " and T____0.ItemCode not in (" & searchTreeFilter & ") "
							End If
							
							sql = sql & " Group By ItmsGrpCod)  "
							
							If myApp.ApplyInvFiltersBy = "I" Then
								sql = sql & " X____1 on X____1.ItmsGrpCod = X____0.ItmsGrpCod "
							End If
							
							sql = sql & " order by 2 "
							set rd = Server.CreateObject("ADODB.RecordSet")
							set rd = conn.execute(sql)
				   		
							 %><select size="1" name="Grupo" style="width: 95%">
											<option></option>
											<% do while not rd.eof %><option value="<%=rd("ItmsGrpCod")%>"><%=rd("ItmsGrpNam")%></option><% rd.movenext
											loop %>
											</select>
						<% Case "FirmCode"

							innerAddStr = ""
							If myApp.GetMinInvBy = "W" Then
								innerAddStr = "inner join OITW T____1 on T____1.ItemCode = T____0.ItemCode and T____1.WhsCode = OLKCommon.dbo.DBOLKGetWhsCode" & Session("ID") & "(" & Session("branch") & ", " & Session("vendid") & ", T____0.ItemCode) "
								fldTbl = "T____1"
							Else
								fldTbl = "T____0"
							End If

							sql = "select X____0.FirmCode, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OMRC', 'FirmName', X____0.FirmCode, X____0.FirmName) FirmName " & _
							"from OMRC X____0 "
							
							If myApp.ApplyInvFiltersBy = "N" Then
								sql = sql & "where X____0.FirmCode in "
							ElseIf myApp.ApplyInvFiltersBy = "I" Then
								sql = sql & " inner join "
							End If
							
							sql = sql & "(select FirmCode from oitm T____0 " & innerAddStr & " where SellItem = 'Y' and " & fldTbl & ".OnHand >= " & myApp.GetMinInv & " "
							
							If QryGroup <> "-1" Then sql = sql & " and QryGroup" & QryGroup & " = 'N' "
							
							If Not IsNull(CatalogFilter) and CatalogFilter <> "" Then
								sql = sql & " and T____0.ItemCode not in (" & CatalogFilter & ") "
							End If
							
							If myApp.GetApplyGenFilter Then
								sql = sql & " and T____0.ItemCode not in (" & myApp.GetGenFilter & ") "
							End If
							
							If Not IsNull(searchTreeFilter) and searchTreeFilter <> "" Then
								sql = sql & " and T____0.ItemCode not in (" & searchTreeFilter & ") "
							End If
							
							sql = sql & " Group By FirmCode) "
							
							If myApp.ApplyInvFiltersBy = "I" Then
								sql = sql & " X____1 on X____1.FirmCode = X____0.FirmCode "
							End If
							
							sql = sql & " order by 2 "
							set rd = Server.CreateObject("ADODB.RecordSet")
							set rd = conn.execute(sql)
							 %>
						<select size="1" name="Marca" style="width: 95%">
						<option></option>
						<% do while not rd.eof %><option value="<%=rd("FirmCode")%>"><%=rd("FirmName")%></option><% rd.movenext
						loop %>
						</select>
						<% Case "Order"
						hdOrder = False %><select size="1" name="<% Select Case ObjID %><% Case 2 %>D3<% Case 4 %>orden1<% End Select %>" style="width: 140px;">
						<% Select Case ObjID
							Case 2 %>
						<option <% If Order1 = "CardType" Then %>selected<% End If %> value="CardType"><%=getadCustomSearchLngStr("DtxtType")%></option>
						<option <% If Order1 = "CardCode" Then %>selected<% End If %> value="CardCode"><%=getadCustomSearchLngStr("DtxtCode")%></option>
						<option <% If Order1 = "CardName" Then %>selected<% End If %> value="CardName"><%=getadCustomSearchLngStr("DtxtName")%></option>
						<option <% If Order1 = "CntctPrsn" Then %>selected<% End If %> value="CntctPrsn"><%=getadCustomSearchLngStr("DtxtContact")%></option>
						<option <% If Order1 = "Balance" Then %>selected<% End If %> value="Balance"><%=getadCustomSearchLngStr("DtxtBalance")%></option>
						<option <% If Order1 = "GroupName" Then %>selected<% End If %> value="GroupName"><%=getadCustomSearchLngStr("DtxtGroup")%></option>
						<option <% If Order1 = "Name" Then %>selected<% End If %> value="Name"><%=getadCustomSearchLngStr("DtxtCountry")%></option>
						<%	Case 4 %>
						<option <% If Order1 = "OITM.ItemCode" Then %>selected<% End If %> value="OITM.ItemCode"><%=getadCustomSearchLngStr("DtxtCode")%></option>
						<option <% If Order1 = "ItemName" Then %>selected<% End If %> value="ItemName"><%=getadCustomSearchLngStr("DtxtDescription")%></option>
						<option <% If Order1 = "Price" Then %>selected<% End If %> value="Price"><%=getadCustomSearchLngStr("DtxtPrice")%></option>
						<% End Select %>
						</select>
						<select size="1" name="<% Select Case ObjID %><% Case 2 %>D5<% Case 4 %>orden2<% End Select %>" style="width: 100px;">
						<option <% If Order2 = "A" Then %>selected<% End If %> value="asc"><%=getadCustomSearchLngStr("DtxtAsc")%></option>
			            <option <% If Order2 = "D" Then %>selected<% End If %> value="desc"><%=getadCustomSearchLngStr("DtxtDesc")%></option>
			            </select>
						<% Case "PriceRange" %>
						<input type="text" name="PriceFrom" size="20" value="" onchange="javascript:<% If userType = "C" Then %>chkThis(this, <%=myApp.MinPrice%>);<% Else %>ChkNum(this);<% End If %>" onkeydown="return chkNumValue(event);" onfocus="javascript:this.select()">
						-
						<input type="text" name="PriceTo" size="20" value="" onchange="javascript:<% If userType = "C" Then %>chkThis(this,null);<% Else %>ChkNum(this);<% End If %>" onkeydown="return chkNumValue(event);" onfocus="javascript:this.select()"></td>
						<% Case "Inventory" %>
						<input type="text" name="InvFrom" style="width: 42%" size="16" value="" onchange="javascript:<% If userType = "C" Then %>chkThis(this,<%=myApp.GetMinInv%>);<% Else %>ChkNum(this);<% End If %>" onkeydown="return chkNumValue(event);" onfocus="javascript:this.select()">
						<% Case "ItemWithImg" %>
						<table border="0" cellpadding="0" cellspacing="1">
							<tr>
								<td width="23">
								<input type="checkbox" class="OptionButton" name="pic" id="pic" value="ON" style="background:background-image"></td>
								<td class="TablasMenutop"><label for="pic"><%=getadCustomSearchLngStr("LtxtItmImg")%></label></td>
							</tr>
						</table>
						<% Case "ItemNew" %>
						<table border="0" cellpadding="0" cellspacing="1">
							<tr>
								<td width="23">
								<input type="checkbox" class="OptionButton" name="new" id="new" value="ON" style="background:background-image"></td>
								<td class="TablasMenutop"><label for="new">
								<%=getadCustomSearchLngStr("LtxtNewItms")%></label></td>
							</tr>
						</table>
						<% Case "ItemProm"
						If optProm and userType = "C" or userType = "V" Then %>
						<table border="0" cellpadding="0" cellspacing="1">
							<tr>
								<td width="23">
								<input name="chkProm" class="OptionButton" id="chkProm" type="checkbox" value="Y" style="background:background-image"></td>
								<td class="TablasMenutop"><label for="chkProm"><%=getadCustomSearchLngStr("LtxtPromOnly")%></label></td>
							</tr>
						</table>
						<% End If
						Case "WishList"
						If optWish and Session("UserName") <> "" Then %>
						<table border="0" cellpadding="0" cellspacing="1">
							<tr>
								<td width="23">
								<input name="chkWL" class="OptionButton" id="chkWL" type="checkbox" value="Y" style="background:background-image"></td>
								<td class="TablasMenutop"><label for="chkWL"><%=getadCustomSearchLngStr("LtxtInWL")%></label></td>
							</tr>
						</table>
						<% End If
						Case "CatType"
						hdCatType = False %>
						<% If 1 = 2 Then %>
						<table border="0" cellpadding="0" cellspacing="1">
							<tr>
									<td width="25">
									<input type="radio" value="T" id="documentT" name="document" class="OptionButton" style="background:background-image;" <% If CatType = "T" Then %>checked<% End If %>></td>
									<td width="63" class="TablasMenutop"><label for="documentT">
									<%=getadCustomSearchLngStr("DtxtStore")%></label></td>
									<td width="24">
									<input type="radio" value="C" id="documentC" name="document" class="OptionButton" style="background:background-image;" <% If CatType = "C" Then %>checked<% End If %>></td>
									<td class="TablasMenutop"><label for="documentC">
									<%=getadCustomSearchLngStr("DtxtCat")%></label></td>
							</tr>
						</table>
						<% End If %>
						<%
						set objViewType = New clsViewType
						objViewType.ID = "document"
						objViewType.Value = CatType
						objViewType.doViewType
						%>
						<% Case "ItemRange" %>
						<input type="text" name="ItemCodeFrom" size="20" onkeydown="return chkMax(event, this, 20);" onchange="javascript:getValue('Itm', this);" onfocus="this.select();"> 
						-
						<input type="text" name="ItemCodeTo" size="20" onkeydown="return chkMax(event, this, 20);" onchange="javascript:getValue('Itm', this);" onfocus="this.select();">
						<% Case "ItmsGrpRange" %>
						<input type="text" name="ItmsGrpNamFrom" size="20" onkeydown="return chkMax(event, this, 20);" onchange="javascript:getValue('ItmGrp', this);" onfocus="this.select();"> 
						-
						<input type="text" name="ItmsGrpNamTo" size="20" onkeydown="return chkMax(event, this, 20);" onchange="javascript:getValue('ItmGrp', this);" onfocus="this.select();">
						<% Case "FirmRange" %>
						<input type="text" name="FirmNameFrom" size="20" onkeydown="return chkMax(event, this, 30);" onchange="javascript:getValue('ItmFrm', this);" onfocus="this.select();"> 
						-
						<input type="text" name="FirmNameTo" size="20" onkeydown="return chkMax(event, this, 30);" onchange="javascript:getValue('ItmFrm', this);" onfocus="this.select();">
						<% Case "InvRange" %>
						<input type="text" name="InvFrom" size="9" onkeydown="return chkNumValue(event);"> 
						- 
						<input type="text" name="InvTo" size="9" onkeydown="return chkNumValue(event);">
						<% Case "CardType" %>
						<% typeCount = 0
						If myAut.HasAuthorization(23) Then typeCount = 1
						If myAut.HasAuthorization(74) Then typeCount = typeCount + 1
						If myAut.HasAuthorization(75) Then typeCount = typeCount + 1 %>
						<select size="1" name="CardType">
						<% If typeCount > 1 Then %><option value=""><%=getadCustomSearchLngStr("DtxtAll")%></option><% End If %>
						<% If myAut.HasAuthorization(23) Then %><option value="C"><% If 1 = 2 Then %>Cliente<% Else %><%=txtClient%><% End If %></option><% End If %>
						<% If myAut.HasAuthorization(74) Then %><option value="S"><%=getadCustomSearchLngStr("DtxtSupplier")%></option><% End If %>
						<% If myAut.HasAuthorization(75) Then %><option value="L"><%=getadCustomSearchLngStr("DtxtLead")%></option><% End If %>
						</select>
						<% Case "BPRange" %>
						<input type="text" name="CardCodeFrom" size="15" onkeydown="return chkMax(event, this, 15);" onchange="javascript:getValue('Crd', this);" onfocus="this.select();">-<input type="text" name="CardCodeTo" size="15" onkeydown="return chkMax(event, this, 15);" onchange="javascript:getValue('Crd', this);" onfocus="this.select();">
						<% Case "BPGrpRange" %>
						<input type="text" name="GroupNameFrom" size="20" onkeydown="return chkMax(event, this, 20);" onchange="javascript:getValue('Grp', this);" onfocus="this.select();">-<input type="text" name="GroupNameTo" size="20" onkeydown="return chkMax(event, this, 20);" onchange="javascript:getValue('Grp', this);" onfocus="this.select();">
						<% Case "BPCntRange" %>
						<input type="text" name="CountryFrom" size="20" onkeydown="return chkMax(event, this, 100);" onchange="javascript:getValue('Cty', this);" onfocus="this.select();">-<input type="text" name="CountryTo" size="20" onkeydown="return chkMax(event, this, 100);" onchange="javascript:getValue('Cty', this);" onfocus="this.select();">
						<% Case "ItmProp", "BPProp" %>
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td width="300">
									<div id="listQryGroupSearch" class="input" style="width: 300px; height: 120px; overflow=auto;">
									<%
									Select Case rs("Variable")
										Case "ItmProp"
											sql = 	"select T0.ItmsTypCod, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITG', 'ItmsGrpNam', T0.ItmsTypCod, T0.ItmsGrpNam) ItmsGrpNam " & _
													"from OITG T0 " & _
													"inner join OLKCustomSearchProp T1 on T1.ObjectCode = 4 and T1.ID = " & adSearchID & " and T1.PropID = T0.ItmsTypCod " & _
													"where T1.Active = 'Y' " & _
													"order by T1.Ordr"
										Case "BPProp"
											sql = 	"select T0.GroupCode, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(1, 'OCQG', 'GroupName', GroupCode, GroupName) GroupName " & _
													"from OCQG T0 " & _
													"inner join OLKCustomSearchProp T1 on T1.ObjectCode = 2 and T1.ID = " & adSearchID & " and T1.PropID = T0.GroupCode " & _
													"where T1.Active = 'Y' " & _
													"order by T1.Ordr"
									End Select
									set rd = conn.execute(sql)
									do while not rd.eof %>
									<div><input class="noborder" type="checkbox" name="chkQryGroup" id="chkQryGroup<%=rd(0)%>" value="<%=rd(0)%>"><label for="chkQryGroup<%=rd(0)%>"><%=rd(1)%></label></div>
									<% rd.movenext
									loop %>
									</div>
								</td>
								<td valign="top">
								<table class="TblGeneral">
									<tr>
										<td><input class="noborder" type="radio" name="QryGroupOp" value="A" id="QryGroupOpA" checked></td>
										<td><label for="QryGroupOpA"><%=getadCustomSearchLngStr("DtxtAnd")%></label></td>
									</tr>
									<tr>
										<td><input class="noborder" type="radio" name="QryGroupOp" value="O" id="QryGroupOpO"></td>
										<td><label for="QryGroupOpO"><%=getadCustomSearchLngStr("DtxtOr")%></label></td>
									</tr>
									<tr>
										<td colspan="2">&nbsp;
									</tr>
									<tr>
										<td><input class="noborder" type="radio" name="QryGroupOp2" value="I" id="QryGroupOp2I" checked></td>
										<td><label for="QryGroupOp2I"><%=getadCustomSearchLngStr("DtxtIn")%></label></td>
									</tr>
									<tr>
										<td><input class="noborder" type="radio" name="QryGroupOp2" value="N" id="QryGroupOp2N"></td>
										<td><label for="QryGroupOp2N"><%=getadCustomSearchLngStr("DtxtNotIn")%></label></td>
									</tr>
								</table>
								</td>
							</tr>
						</table>
				   <%	End Select
					End Select %></td>
				</tr>
				<% rs.movenext
				loop %>

			<% If Session("RetVal") = "" and ObjID = 4 Then %>

					<tr>
						<td class="TablasNoticias"><%=getadCustomSearchLngStr("LtxtPriceList")%></td>
						<td class="TblGeneral">
						<select size="1" name="CPList">
						<option value=""><%=getadCustomSearchLngStr("LtxtWithoutPrice")%></option>
						<option value="X"><%=getadCustomSearchLngStr("LtxtDocPrice")%></option>
						<% 
						
						set cmd = Server.CreateObject("ADODB.Command")
						cmd.ActiveConnection = connCommon
						cmd.CommandType = &H0004
						cmd.CommandText = "DBOLKGetPriceListFiltered" & Session("ID")
						cmd.Parameters.Refresh()
						cmd("@LanID") = Session("LanID")
						cmd("@UserAccess") = Session("UserAccess")
						cmd("@SlpCode") = Session("vendid")
						set rs = cmd.execute()
						Do While NOT RS.EOF %>
						<option <% If CStr(Request("CPList")) = CStr(rs("ListNum")) then response.write "selected" %> value="<%=RS("Listnum")%>"><%=myHTMLEncode(RS("ListName"))%></option>
						<% RS.MoveNext
						loop %>
						</select></td>
					</tr>
					<tr>
						<td class="TablasNoticias"><%=getadCustomSearchLngStr("LtxtDoc")%></td>
						<td class="TblGeneral">
						<select size="1" name="sourceDoc">
						<option value=""><%=getadCustomSearchLngStr("DtxtCat")%></option>
						<optgroup label="<%=getadCustomSearchLngStr("LtxtSale")%>">
							<option value="23"><% If 1 = 2 Then %>Cotización<% Else %><%=txtQuote%><% End If %></option>
							<option value="17"><% If 1 = 2 Then %>Pedido<% Else %><%=txtOrdr%><% End If %></option>
							<option value="15"><% If 1 = 2 Then %>Entregas<% Else %><%=txtOdlns%><% End If %></option>
							<option value="13"><% If 1 = 2 Then %>Factura<% Else %><%=txtInv%><% End If %></option>
						</optgroup>
						<optgroup label="<%=getadCustomSearchLngStr("LtxtPur")%>">
							<option value="22"><% If 1 = 2 Then %>Orden 
							de Compra<% Else %><%=txtOpor%><% End If %></option>
							<option value="20"><% If 1 = 2 Then %>Entrada de Mercancia OP<% Else %><%=txtOpdn%><% End If %></option>
							<option value="18"><% If 1 = 2 Then %>Comp. 
							de Compra<% Else %><%=txtOpch%><% End If %></option>
						</optgroup>
						<optgroup label="<%=getadCustomSearchLngStr("DtxtOLK")%>">
							<option value="-4"><%=getadCustomSearchLngStr("DtxtLogNum")%></option>
						</optgroup>
						</select></td>
					</tr>
					<tr>
						<td class="TablasNoticias"><%=getadCustomSearchLngStr("LtxtDocNum")%></td>
						<td class="TblGeneral">
						<input type="text" name="DocNum" size="9" onchange="javascript:ChkNum(this);"></td>
					</tr>
			<% End If %>
			</table>
			</td>
		</tr>
		<tr class="TablasNoticias">
			<td>
			<p align="center">
			<input type="submit" name="btnSearch" value="<%=getadCustomSearchLngStr("DtxtSearch")%>"></td>
		</tr>
	</table>
</div>
<% If hdCatType Then %><input type="hidden" name="document" value="<%=CatType%>"><% End If %>
<% If hdOrder Then %><input type="hidden" name="orden1" value="<%=Order1%>"><input type="hidden" name="<%=Order2%>" value="asc"><% End If %>
<input type="hidden" name="ID" value="<%=adSearchID%>">
<input type="hidden" name="adObjID" value="<%=ObjID%>">
<input type="hidden" name="adSearch" value="Y">
<input type="hidden" name="cmd" value="<% Select Case ObjID %><% Case 2 %>clientsSearch<% Case 4 %><% If Session("RetVal") <> "" Then %>searchCart<% Else %>searchCatalog<% End If %><% End Select %>">
<input type="hidden" name="isSubmit" value="N">
</form>
<iframe id="ifGetValue" name="ifGetValue" style="display: none" height="99" width="256" src=""></iframe>
<form method="post" target="ifGetValue" name="frmGetValue" action="topGetValue.asp">
<input type="hidden" name="Type" value="">
<input type="hidden" name="searchStr" value="">
</form>
<script language="javascript">
var txtValNumVal = '<%=getadCustomSearchLngStr("DtxtValNumVal")%>';
</script>
<script language="javascript" src="adCustomSearch.js"></script>
<script type="text/javascript">
function doLoadDesc()
{
	<% If notNullVars <> "" Then 
	ArrVal = Split(notNullVars, ", ")
	for i = 0 to UBound(ArrVal)
	ArrVal2 = Split(ArrVal(i),"~")
	If ArrVal2(2) <> "CL" Then %>
	if (document.frmVars.<%=ArrVal2(0)%>.value == '')
	{
		alert('<%=getadCustomSearchLngStr("LtxtEnterFld")%>'.replace('{0}', '<%=Replace(ArrVal2(1), "'", "\'")%>'));
		document.frmVars.<%=ArrVal2(0)%>.focus();
		return false;
	}
	<% Else %>
	if (!isChkboxChecked(document.frmVars.<%=ArrVal2(0)%>))
	{
		alert('<%=getadCustomSearchLngStr("LtxtEnterFld")%>'.replace('{0}', '<%=Replace(ArrVal2(1), "'", "\'")%>'));
		return false;
	}
<%	End If 
	Next
	End If %>
	<% If Session("RetVal") = "" and ObjID = 4 Then %>
	if (document.frmVars.sourceDoc.value != '' && (!MyIsNumeric(document.frmVars.DocNum.value) || document.frmVars.DocNum.value == ''))
	{
		alert('<%=getadCustomSearchLngStr("LtxtValDocCatNum")%>');
		document.frmSmallSearch.DocNum.focus();
		return false;
	}
	else if (document.frmVars.CPList.value == 'X' && document.frmVars.sourceDoc.value == '')
	{
		alert('<%=getadCustomSearchLngStr("LtxtValDocDocType")%>');
		return false;
	}
	<% End If %>
	return true;
}

	<% 
	If rs.RecordCount > 0 Then
	rs.movefirst
	do while not rs.eof
	If rs("Type") = "DP" Then %>
    Calendar.setup({
        inputField     :    "var<%=rs("varID")%>",     // id of the input field
        ifFormat       :    "<%=GetCalendarFormatString%>",      // format of the input field
        button         :    "btn<%=rs("varID")%>",  // trigger for the calendar (button ID)
        align          :    "Bl",           // alignment (defaults to "Bl")
        singleClick    :    true
    });
    <% End If
    rs.movenext
    loop
    End If %>
</script>
<script language="javascript">
<% If Request.Form("isSubmit") = "" Then %>document.body.onload += reload('');<% End If %>
</script>
<% set rd = nothing
set rs = nothing %><%
Function getCustomSearchValsCSQL(doQuery, Qry)
	If doQuery Then
		retVal= " "
	Else
		retVal = ""
	End If
	rBase.Filter = "varID = " & rs("varID")
	do while not rBase.eof
		If Request.Form("var" & rBase("BaseID")) <> "" Then
			If doQuery Then
				If rBase("DataType") = "nvarchar" Then 
					MaxVar = "(" & rBase("MaxChar") & ")"
				ElseIf rBase("DataType") = "numeric" Then
					MaxVar = "(19,6)"
				Else
					MaxVar = ""
				End If
				retVal = retVal & "declare @" & rBase("Variable") & " " & rBase("DataType") & " " & MaxVar & " "
				Select Case rBase("DataType") 
					Case "nvarchar" 
						retVal = retVal & "set @" & rBase("Variable") & " = N'" & saveHTMLDecode(Request.Form("var" & rBase("BaseID")), False) & "' "
					Case "datetime"
						retVal = retVal & "set @" & rBase("Variable") & " = Convert(datetime,'" & SaveSqlDate(Request.Form("var" & rBase("BaseID"))) & "',120) "
					Case Else
						retVal = retVal & "set @" & rBase("Variable") & " = " & Request.Form("var" & rBase("BaseID")) & " "
				End Select
			Else
				retVal = retVal & "&var" & rBase("BaseID") & "=" & Request.Form("var" & rBase("BaseID"))
			End If
		Else
			selectRepValsCName = rBase("Name")
			enableControl = False
			Exit Do
		End If
	rBase.movenext
	loop
	If doQuery Then
		retVal = retVal & "declare @LanID int set @LanID = " & Session("LanID") & " "
	End If
	If doQuery Then retVal = retVal & Qry
	getCustomSearchValsCSQL = retVal
End Function %>
<% Select Case userType
Case "C" %><!--#include file="clientBottom.asp"-->
<% Case "V" %><!--#include file="agentBottom.asp"-->
<% End Select %>
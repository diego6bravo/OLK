<% addLngPathStr = "" %>
<!--#include file="lang/adSearch.asp" -->
<%  
ID = Request("ID")
ObjID = Request("adObjID")

Select Case ObjID
	Case 2
		SearchCmd = "searchresult"
	Case 4
		SearchCmd = "searchItems"
End Select

sql = 	"select Count('')  " & _
		"from OLKCustomSearch T0 " & _
		"where T0.ObjectCode = " & ObjID & " and T0.Status = 'Y' and exists(select '' from OLKCustomSearchSession where ObjectCode = T0.ObjectCode and ID = T0.ID and SessionID = 'P') "
set rs = conn.execute(sql)
MultiAdSearch = rs(0) > 1


sql = "select IsNull(T2.AlterName, T1.Name) Name, T1.Order1, T1.Order2 " & _
	"from OLKCommon T0 " & _
	"inner join OLKCustomSearch T1 on T1.ObjectCode = " & ObjID & " and T1.ID = " & ID & " " & _
	"left outer join OLKCustomSearchAlterNames T2 on T2.ObjectCode = T1.ObjectCode and T2.ID = T1.ID and T2.LanID = " & Session("LanID")
set rs = conn.execute(sql)
AdSearchName = rs("Name")
If Request("orden1") = "" Then
	Order1 = rs("Order1")
	Order2 = rs("Order2")
Else
	Order1 = Request("orden1")
	Order2 = UCase(Left(Request("orden2"), 1))
End If	

	
If Order1 = "" or IsNull(Order1) Then
	Select Case myApp.GetDefCatOrdr
		Case "C"
			Order1 = "OITM.ItemCode"
		Case "N"
			Order1 = "ItemName"
	End Select
End If

rs.close %>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
    <tr>
      <td bgcolor="#9BC4FF">
      <table border="0" cellpadding="0" bordercolor="#111111" width="100%">
        <!--#include file="C_Art/CardNameAdd.asp" -->
        <tr>
          <td width="100%">
          <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><%=dSearchName%></font></b></td>
        </tr>
        <tr>
          <td width="100%"><font face="Verdana" size="1"><%=AdSearchName%></font></td>
        </tr>
        <% If Request("getValType") = "" or Request("getValType") <> "" and InStr(Request("getValVal"), "*") = 0 Then
        
        setToVal = ""
		ItemCodeFrom 	= ListAdSearchGetValue("ItemCodeFrom")
		ItemCodeTo		= ListAdSearchGetValue("ItemCodeTo")
		ItmsGrpNamFrom	= ListAdSearchGetValue("ItmsGrpNamFrom")
		ItmsGrpNamTo	= ListAdSearchGetValue("ItmsGrpNamTo")
		FirmNameFrom	= ListAdSearchGetValue("FirmNameFrom")
		FirmNameTo		= ListAdSearchGetValue("FirmNameTo")
		CardCodeFrom	= ListAdSearchGetValue("CardCodeFrom")
		CardCodeTo		= ListAdSearchGetValue("CardCodeTo")
		GroupNameFrom	= ListAdSearchGetValue("GroupNameFrom")
		GroupNameTo		= ListAdSearchGetValue("GroupNameTo")
		CountryFrom		= ListAdSearchGetValue("CountryFrom")
		CountryTo		= ListAdSearchGetValue("CountryTo")
		
		Function ListAdSearchGetValue(Fld)
			If Request("getValType") = "" or Request("getValType") <> "" and Request("getValFld") <> Fld Then
				If setToVal <> "" and Request(Fld) = "" Then
					ListAdSearchGetValue = setToVal
				Else
					ListAdSearchGetValue = Request(Fld)
				End If
				setToVal = ""
			Else
				
				Dim getVal
				set getVal = new clsGetValValue
				getVal.ValueType = Request("getValType")
				getVal.ValueField = Request("getValFld")
				getVal.Value = Request("getValVal")
				
				NewValue = getVal.GetValue
				
				If NewValue <> "" Then
					ListAdSearchGetValue = NewValue
					
					If Right(Fld, 4) = "From" Then setToVal = NewValue
				Else
					ListAdSearchGetValue = ""
				End If 
			End If
		End Function

		 %>
		<form method="POST" name="frmVars" action="operaciones.asp">
		<input type="hidden" name="getValType" value="">
		<input type="hidden" name="getValFld" value="">
		<input type="hidden" name="getValVal" value="">
		<input type="hidden" name="adSearch" value="Y">
        <tr>
          <td width="100%">
			<table border="0" cellspacing="0" width="100">
			<%
			hdOrder = True
			sql = "select T0.VarID, IsNull(T1.alterName, T0.Name) Name, T0.Variable, T0.Type, T0.DataType, T0.Query, T0.QueryField, T0.MaxChar, T0.NotNull, T0.DefVars, T0.DefValBy, T0.DefValValue, T0.DefValDate, " & _
				"Case When Exists(select 'A' from OLKCustomSearchVarsBase where ObjectCode = T0.ObjectCode and ID = T0.ID and BaseID = T0.VarID) Then 'Y' Else 'N' End IsBase, " & _
				"OLKCommon.dbo.DBOLKGetCustomSearchVarTarget" & Session("ID") & "(T0.ObjectCode, T0.ID, T0.VarID) TargetIndex " & _
				"from OLKCustomSearchVars T0 " & _
				"left outer join OLKCustomSearchVarsAlterNames T1 on T1.ObjectCode = T0.ObjectCode and T1.ID = T0.ID and T1.VarID = T0.VarID and T1.LanID = " & Session("LanID") & " " & _
				"where T0.ObjectCode = " & ObjID & " and T0.ID = " & ID & " and ([Type] <> 'S' or [Type] = 'S' and T0.Variable <> 'CatType') order by T0.Ordr asc"
			rs.open sql, conn, 3, 1
			set rQVal = Server.CreateObject("ADODB.RecordSet")
			set rBase = Server.CreateObject("ADODB.RecordSet")
			sql = "select T1.VarID, T1.BaseID, Variable, IsNull(alterName, Name) Name, DataType, MaxChar " & _
					"from OLKCustomSearchVars T0 " & _
					"inner join OLKCustomSearchVarsBase T1 on T1.ObjectCode = T0.ObjectCode and T1.ID = T0.ID and T1.BaseID = T0.VarID " & _
					"left outer join OLKCustomSearchVarsAlterNames T2 on T2.ObjectCode = T0.ObjectCode and T2.ID = T0.ID and T2.VarID = T0.VarID and T2.LanID = " & Session("LanID") & " " & _
					"where T0.ObjectCode = " & ObjID & " and T0.ID = " & ID & " "
			rBase.open sql, conn, 3, 1
			do while not rs.eof
				enableControl = True
				If rs("NotNull") = "Y" Then
					If notNullVars <> "" Then notNullVars = notNullVars & ", "
					notNullVars = notNullVars & "var" & rs("VarID") & "~" & rs("Name")
				End If
				Select Case rs("Type")
					Case "DD", "L", "CL"
					   If rs("DefVars") = "F" Then
					   		sql = "select Value, Description " & _
					   				"from OLKCustomSearchVarsVals " & _
					   				"where ObjectCode = " & ObjID & " and ID = " & Request("ID") & " and VarID = " & rs("VarID")
					   Else
							sql = getAdSearchSQL(true, rs("Query"))
					   End If
					'Case "L"
					 '  If rs("DefVars") = "F" Then
					 '  		sql = "select valValue, valText from OLKCustomSearchVarsVals where and ObjectCode = " & ObjID & " and ID = " & Request("ID") & " and VarID = " & rs("VarID")
					 '  Else
					 '  		sql = getAdSearchSQL(true, rs("Query"))
					 '  End If
					Case "Q"
						If rs("DefVars") = "Q" Then
							sql = getAdSearchSQL(false, "")
						End If
				End Select
				If Request("isSubmit") <> "R" Then
					If rs("Type") <> "DP" and rs("DefValBy") = "V" Then
						defValue = rs("DefValValue")
					ElseIf rs("Type") = "DP" and rs("DefValBy") = "V" Then
						defValue = FormatDate(rs("DefValDate"), False)
					ElseIf rs("DefValBy") = "Q" Then
						sqlVal = getAdSearchSQL(true, rs("DefValValue"))
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
					defValue = Request("var" & rs("VarID"))
					If rs("Type") = "CL" Then
						defValueDesc = Request("var" & rs("VarID") & "Desc")
					End If
				End If %>
			<tr class="TblAfueraMnu">
				<td valign="top"><font face="Verdana" size="1"><% Select Case rs("Type")
					Case "S"
						Select Case rs("Variable")
							Case "Search" 
								Response.Write getadSearchLngStr("DtxtSearch")
							Case "ItmsGrpCod", "ItmsGrpRange"
								Response.Write txtAlterGrp
								If rs("Variable") = "ItmsGrpRange" Then Response.Write strFromTo
							Case "FirmCode", "FirmRange"
								Response.Write txtAlterFrm
								If rs("Variable") = "FirmRange" Then Response.Write strFromTo
							Case "Order" 
								Response.Write getadSearchLngStr("DtxtOrder")
							Case "PriceRange" 
								Response.Write getadSearchLngStr("DtxtPrice") & strFromTo
							Case "Inventory"
								Response.Write getadSearchLngStr("LtxtInvMoreThen") 
							Case "InvRange"
								Response.Write getadSearchLngStr("LtxtInventory") & strFromTo
							Case "CatType"
								Response.Write getadSearchLngStr("LtxtCatalogType")
							Case "ItemRange"
								Response.Write getadSearchLngStr("DtxtItem") & strFromTo
							Case "CardType"
								Response.Write getadSearchLngStr("DtxtType")
							Case "BPRange"
								Response.Write getadSearchLngStr("DtxtBPCode") & strFromTo
							Case "BPGrpRange"
								Response.Write getadSearchLngStr("DtxtGroup") & strFromTo
							Case "BPCntRange"
								Response.Write getadSearchLngStr("DtxtCountry") & strFromTo
							Case "ItmProp", "BPProp"
								Response.Write getadSearchLngStr("DtxtProp")					
							Case Else
								Response.Write "&nbsp;"
						End Select
					Case Else 
						Response.Write Server.HTMLEncode(rs("Name"))
					End Select%>&nbsp;</font></td>
				<td width="16"><% Select Case rs("Type") %>
					<% Case "Q" %><% If enableControl Then %><a href="javascript:doQuery(<%=rs("VarID")%>);"><% End If %><img border="0" src="images/<%=Session("rtl")%>flechaselec.gif"><% If enableControl Then %></a><% End If %>
				<% Case "DP" %><a href="javascript:doCal(<%=rs("VarID")%>);"><img border="0" src="images/cal.gif" width="16" height="16"></a>
				<% End Select %></td>
				<td>
				<% Select Case rs("Type") %>
				<% Case "DD"
				If enableControl Then set rd = conn.execute(sql) %>
				<select name="var<%=rs("VarID")%>" size="1"  <% If Not enableControl Then %>disabled<% End If %> <% If rs("IsBase") = "Y" Then %> onchange="reload(<%=rs("targetIndex")%>);"<% End If %>>
				<% If enableControl Then %>
				<option></option>
				<% do while not rd.eof %><option <% If defValue = CStr(rd(0)) Then %>selected<% End If %> value="<%=rd(0)%>"><%=myHTMLEncode(rd(1))%></option>
				<% rd.movenext
				loop
				Else %>
				<option value=""><%=getadSearchLngStr("LtxtSelectEnter")%> "<%=selectName%>"</option>
				<% End If %>
				</select>
				<% Case "T" %><input name="var<%=rs("VarID")%>" type="text" onchange="chkNum(this, '<%=rs("DataType")%>');<% If rs("IsBase") = "Y" Then %>reload(<%=rs("targetIndex")%>);<% End If %>" value="<%=defValue%>" size="15">
				<% Case "L"
				If enableControl Then set rd = conn.execute(sql) %>
				<select name="var<%=rs("VarID")%>" size="5" <% If Not enableControl Then %>disabled<% End If %> <% If rs("IsBase") = "Y" Then %> onchange="reload(<%=rs("targetIndex")%>);"<% End If %>><%
				If enableControl Then
				do while not rd.eof
				%><option <% If defValue = CStr(rd(0)) Then %>selected<% End If %> value="<%=rd(0)%>"><%=myHTMLEncode(rd(1))%></option>
				<% rd.movenext
				loop
				Else %>
				<option value=""><%=getadSearchLngStr("LtxtSelectEnter")%> "<%=selectName%>"</option>
				<% End If %>
				</select>
				<% Case "Q" %><input name="var<%=rs("VarID")%>" type="text" readonly size="15" <% If enableControl Then %> value="<%=defValue%>" onclick="javascript:doQuery(<%=rs("VarID")%>);"<% Else %> disabled value="Seleccione/Introduzca &quot;<%=selectName%>&quot;" <% End If %> <% If rs("IsBase") = "Y" Then %> onchange="reload(<%=rs("targetIndex")%>);"<% End If %>>
				<% Case "DP" %><input name="var<%=rs("VarID")%>" type="text" readonly size="12" onclick="doCal(<%=rs("VarID")%>);" value="<%=defValue%>" <% If rs("IsBase") = "Y" Then %> onchange="reload(<%=rs("targetIndex")%>);"<% End If %>>
				<% Case "CL" %><input name="var<%=rs("VarID")%>Desc" type="text" readonly size="15" <% If enableControl Then %> value="<%=defValueDesc%>" onclick="javascript:doChkList(<%=rs("VarID")%>);"<% Else %> disabled value="Seleccione/Introduzca &quot;<%=selectName%>&quot;" <% End If %> <% If rs("IsBase") = "Y" Then %> onchange="reload(<%=rs("targetIndex")%>);"<% End If %>>
				<input name="var<%=rs("VarID")%>" type="hidden" <% If enableControl Then %> value="<%=defValue%>" <% End If %>>
				<% Case "S"
				   		Select Case rs("Variable")
				   			Case "Search" %><table cellpadding="0" cellspacing="0" border="0">
				   				<tr>
				   					<td><input type="text" name="string" value="<%=Request("string")%>" size="25" style="font-size:12px;"></td>
				   				</tr><% If ObjID = 4 Then %>
				   				<% If myApp.SearchExactP Then %>
				   				<tr>
				   					<td>
									<p align="center">
									<font face="Verdana" size="1">
									<input type="radio" value="E" name="rdSearchAs" id="rdSearchAsE" <% If Request("rdSearchAs") = "" and myApp.SearchMethodP = "E" or Request("rdSearchAs") = "E" Then %>checked<% End If %>><label for="rdSearchAsE"><%=getadSearchLngStr("LtxtExact")%></label><input type="radio" name="rdSearchAs" id="rdSearchAsS" value="S" <% If Request("rdSearchAs") = "" and myApp.SearchMethodP = "L" or Request("rdSearchAs") = "S" Then %>checked<% End If %>><label for="rdSearchAsS"><%=getadSearchLngStr("LtxtLike")%></label></font></td>
				   				</tr>
				   				<% Else %>
				   				<input type="hidden" name="rdSearchAs" value="S">
				   				<% End If %>
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
							
							If myApp.CarArt <> "-1" Then sql = sql & " and QryGroup" & myApp.CarArt & " = 'N' "
							
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
											<% do while not rd.eof %><option <% If Request("Grupo") = rd("ItmsGrpCod") Then %>selected<% End If %> value="<%=rd("ItmsGrpCod")%>"><%=rd("ItmsGrpNam")%></option><% rd.movenext
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
							
							If myApp.CarArt <> "-1" Then sql = sql & " and QryGroup" & myApp.CarArt & " = 'N' "
							
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
						<% do while not rd.eof %><option <% If Request("Marca") = rd("FirmCode") Then %>selected<% End If %> value="<%=rd("FirmCode")%>"><%=rd("FirmName")%></option><% rd.movenext
						loop %>
						</select>
						<% Case "Order"
						hdOrder = False %><select size="1" name="orden1" style="width: 140px;">
						<% Select Case ObjID
							Case 2 %>
						<option <% If Order1 = "CardType" Then %>selected<% End If %> value="CardType"><%=getadSearchLngStr("DtxtType")%></option>
						<option <% If Order1 = "CardCode" Then %>selected<% End If %> value="CardCode"><%=getadSearchLngStr("DtxtCode")%></option>
						<option <% If Order1 = "CardName" Then %>selected<% End If %> value="CardName"><%=getadSearchLngStr("DtxtName")%></option>
						<option <% If Order1 = "CntctPrsn" Then %>selected<% End If %> value="CntctPrsn"><%=getadSearchLngStr("DtxtContact")%></option>
						<option <% If Order1 = "Balance" Then %>selected<% End If %> value="Balance"><%=getadSearchLngStr("DtxtBalance")%></option>
						<option <% If Order1 = "GroupName" Then %>selected<% End If %> value="GroupName"><%=getadSearchLngStr("DtxtGroup")%></option>
						<option <% If Order1 = "Name" Then %>selected<% End If %> value="Name"><%=getadSearchLngStr("DtxtCountry")%></option>
						<% 	Case 4 %>
						<option <% If Order1 = "OITM.ItemCode" Then %>selected<% End If %> value="OITM.ItemCode"><%=getadSearchLngStr("DtxtCode")%></option>
						<option <% If Order1 = "ItemName" Then %>selected<% End If %> value="ItemName"><%=getadSearchLngStr("DtxtDescription")%></option>
						<option <% If Order1 = "Price" Then %>selected<% End If %> value="Price"><%=getadSearchLngStr("DtxtPrice")%></option>
						<% End Select %>
						</select>
						<select size="1" name="orden2" style="width: 100px;">
						<option <% If Order2 = "A" Then %>selected<% End If %> value="asc"><%=getadSearchLngStr("DtxtAsc")%></option>
			            <option <% If Order2 = "D" Then %>selected<% End If %> value="desc"><%=getadSearchLngStr("DtxtDesc")%></option>
			            </select>
						<% Case "PriceRange" %>
						<input type="number" min="0" step="<%=GetNumberStep(myApp.PriceDec)%>" name="PriceFrom" value="<%=PriceFrom%>" size="11" style="width: 42%" value="" onchange="javascript:<% If userType = "C" Then %>chkThis(this, <%=myApp.MinPrice%>);<% Else %>ChkNum(this);<% End If %>">
						<input type="number" min="0" step="<%=GetNumberStep(myApp.PriceDec)%>" name="PriceTo" value="<%=PriceTo%>" size="11" style="width: 42%" value="" onchange="javascript:<% If userType = "C" Then %>chkThis(this,null);<% Else %>ChkNum(this);<% End If %>"></td>
						<% Case "Inventory" %>
						<input type="number" name="InvFrom" min="0" step="<%=GetNumberStep(myApp.QtyDec)%>" style="width: 42%" size="16" value="" onchange="javascript:<% If userType = "C" Then %>chkThis(this,<%=myApp.GetMinInv%>);<% Else %>ChkNum(this);<% End If %>">
						<% Case "ItemWithImg" %>
						<table border="0" cellpadding="0" cellspacing="1" width="100%">
							<tr>
								<td width="23">
								<input type="checkbox" class="OptionButton" name="pic" <% If Request("pic") = "ON" Then %>checked<% End If %> id="pic" value="ON" style="background:background-image"></td>
								<td><font face="Verdana" size="1"><label for="pic"><%=getadSearchLngStr("LtxtItmImg")%></label></font></td>
							</tr>
						</table>
						<% Case "ItemNew" %>
						<table border="0" cellpadding="0" cellspacing="1" width="100%">
							<tr>
								<td width="23">
								<input type="checkbox" class="OptionButton" name="new" <% If Request("new") = "ON" Then %>checked<% End If %> id="new" value="ON" style="background:background-image"></td>
								<td><font face="Verdana" size="1"><label for="new">
								<%=getadSearchLngStr("LtxtNewItms")%></label></font></td>
							</tr>
						</table>
						<% Case "ItemProm" %>
						<table border="0" cellpadding="0" cellspacing="1" width="100%">
							<tr>
								<td width="23">
								<input name="chkProm" class="OptionButton" id="chkProm" <% If Request("chkProm") = "Y" Then %>checked<% End If %> type="checkbox" value="Y" style="background:background-image"></td>
								<td><font face="Verdana" size="1"><label for="chkProm"><%=getadSearchLngStr("LtxtPromOnly")%></label></font></td>
							</tr>
						</table>
						<% Case "WishList" %>
						<table border="0" cellpadding="0" cellspacing="1" width="100%">
							<tr>
								<td width="23">
								<input name="chkWL" class="OptionButton" id="chkWL"" <% If Request("chkWL") = "Y" Then %>checked<% End If %> type="checkbox" value="Y" style="background:background-image"></td>
								<td><font face="Verdana" size="1"><label for="chkWL"><%=getadSearchLngStr("LtxtInWL")%></label></font></td>
							</tr>
						</table>
						<% Case "ItemRange" %>
						<input type="text" name="ItemCodeFrom" value="<%=ItemCodeFrom%>" size="20" maxlength="20" onchange="javascript:getValue('Itm', this);" onfocus="this.select();" onmouseup="event.preventDefault()">
						<input type="text" name="ItemCodeTo" value="<%=ItemCodeTo%>" size="20" maxlength="20" onchange="javascript:getValue('Itm', this);" onfocus="this.select();" onmouseup="event.preventDefault()">
						<% Case "ItmsGrpRange" %>
						<input type="text" name="ItmsGrpNamFrom" value="<%=ItmsGrpNamFrom%>" size="20" maxlength="20" onchange="javascript:getValue('ItmGrp', this);" onfocus="this.select();" onmouseup="event.preventDefault()">
						<input type="text" name="ItmsGrpNamTo" value="<%=ItmsGrpNamTo%>" size="20" maxlength="20" onchange="javascript:getValue('ItmGrp', this);" onfocus="this.select();" onmouseup="event.preventDefault()">
						<% Case "FirmRange" %>
						<input type="text" name="FirmNameFrom" value="<%=FirmNameFrom%>" size="20" maxlength="30" onchange="javascript:getValue('ItmFrm', this);" onfocus="this.select();" onmouseup="event.preventDefault()">
						<input type="text" name="FirmNameTo" value="<%=FirmNameTo%>" size="20" maxlength="30" onchange="javascript:getValue('ItmFrm', this);" onfocus="this.select();" onmouseup="event.preventDefault()">
						<% Case "InvRange" %>
						<input type="number" min="0" step="<%=GetNumberStep(myApp.QtyDec)%>" onclick="this.selectionStart=0;this.selectionEnd=this.value.length;" name="InvFrom" value="<%=Request("InvFrom")%>" size="9" onchange="javascript:ChkNum(this);"> 
						<input type="number" min="0" step="<%=GetNumberStep(myApp.QtyDec)%>" onclick="this.selectionStart=0;this.selectionEnd=this.value.length;" name="InvTo" value="<%=Request("InvTo")%>" size="9" onchange="javascript:ChkNum(this);">
						<% Case "CardType" %>
						<% typeCount = 0
						If myAut.HasAuthorization(23) Then typeCount = 1
						If myAut.HasAuthorization(75) Then typeCount = typeCount + 1 %>
						<select size="1" name="CardType">
						<% If typeCount > 1 Then %><option value=""><%=getadSearchLngStr("DtxtAll")%></option><% End If %>
						<% If myAut.HasAuthorization(23) Then %><option value="C"><% If 1 = 2 Then %>Cliente<% Else %><%=txtClient%><% End If %></option><% End If %>
						<% If myAut.HasAuthorization(75) Then %><option value="L"><%=getadSearchLngStr("DtxtLead")%></option><% End If %>
						</select>
						<% Case "BPRange" %>
						<input type="text" name="CardCodeFrom" value="<%=CardCodeFrom%>" size="20" maxlength="30" onchange="javascript:getValue('Crd', this);" onfocus="this.select();" onmouseup="event.preventDefault()">
						<input type="text" name="CardCodeTo" value="<%=CardCodeTo%>" size="20" maxlength="30" onchange="javascript:getValue('Crd', this);" onfocus="this.select();" onmouseup="event.preventDefault()">
						<% Case "BPGrpRange" %>
						<input type="text" name="GroupNameFrom" value="<%=GroupNameFrom%>" size="20" maxlength="30" onchange="javascript:getValue('Grp', this);" onfocus="this.select();" onmouseup="event.preventDefault()">
						<input type="text" name="GroupNameTo" value="<%=GroupNameTo%>" size="20" maxlength="30" onchange="javascript:getValue('Grp', this);" onfocus="this.select();" onmouseup="event.preventDefault()">
						<% Case "BPCntRange" %>
						<input type="text" name="CountryFrom" value="<%=CountryFrom%>" size="20" maxlength="30" onchange="javascript:getValue('Cty', this);" onfocus="this.select();" onmouseup="event.preventDefault()">
						<input type="text" name="CountryTo" value="<%=CountryTo%>" size="20" maxlength="30" onchange="javascript:getValue('Cty', this);" onfocus="this.select();" onmouseup="event.preventDefault()">
						<% Case "ItmProp", "BPProp" %>
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td valign="top">
								<table cellpadding="0" cellspacing="0" border="0">
									<tr>
										<td><textarea name="chkQryGroupDesc" rows="4" readonly cols="18" onkeydown="return false;" onclick="doPropList();"><%=Request("chkQryGroupDesc")%></textarea>
										<input type="hidden" name="chkQryGroup" value="<%=Request("chkQryGroup")%>"></td>
										<td valign="bottom"><input type="button" name="btnQryGroup" value="..." onclick="doPropList();"></td>
									</tr>
								</table>
								</td>
							</tr>
							<tr>
								<td valign="top">
								<table class="TblGeneral">
									<tr>
										<td><input class="noborder" type="radio" name="QryGroupOp" value="A" id="QryGroupOpA" checked></td>
										<td><font face="Verdana" size="1"><label for="QryGroupOpA"><%=getadSearchLngStr("DtxtAnd")%></label></font></td>
										<td><input class="noborder" type="radio" name="QryGroupOp" value="O" id="QryGroupOpO"></td>
										<td><font face="Verdana" size="1"><label for="QryGroupOpO"><%=getadSearchLngStr("DtxtOr")%></label></font></td>
									</tr>
									<tr>
										<td><input class="noborder" type="radio" name="QryGroupOp2" value="I" id="QryGroupOp2I" checked></td>
										<td><font face="Verdana" size="1"><label for="QryGroupOp2I"><%=getadSearchLngStr("DtxtIn")%></label></font></td>
										<td><input class="noborder" type="radio" name="QryGroupOp2" value="N" id="QryGroupOp2N"></td>
										<td><font face="Verdana" size="1"><label for="QryGroupOp2N"><%=getadSearchLngStr("DtxtNotIn")%></label></font></td>
									</tr>
								</table>
								</td>
							</tr>
						</table>
				   <%	End Select
				End Select %>
				</td>
			</tr>
			<% rs.movenext
			loop %>
			</table>
          </td>
        </tr>
		<tr class="TblAfueraMnu">
			<td colspan="3">
			<p align="center">
			<input type="submit" value="<%=getadSearchLngStr("DtxtSearch")%>" name="btnSearch" onclick="return valFrm();"></td>
		</tr>
		<input type="hidden" name="ID" value="<%=Request("ID")%>">
		<input type="hidden" name="adObjID" value="<%=ObjID%>">
		<% If hdOrder Then %><input type="hidden" name="orden1" value="<%=Order1%>"><input type="hidden" name="<%=Order2%>" value="asc"><% End If %>
		<input type="hidden" name="editVar" value="">
		<input type="hidden" name="cmd" value="<%=SearchCmd%>">
		<input type="hidden" name="slist" value="<%=Request("slist")%>">
		<input type="hidden" name="isSubmit" value="Y">
    	</form>
    	<% Else %>
		<tr>
			<td colspan="3"><b><font face="Verdana" size="1"><%=getadSearchLngStr("LtxtSelVal")%></font></b></td>
		</tr>
		<tr>
			<td colspan="3">
			<% 
			set getValSelect = New clsGetValValueSelect
			getValSelect.ValueType = Request("getValType")
			getValSelect.ValueField = Request("getValFld")
			getValSelect.Value = Request("getValVal")
			getValSelect.OnClick = "javascript:setSmallSearchVal('{0}');"
			getValSelect.OnCancel = "cancelSmallSearchVal();"
			getValSelect.ShowValues
			%>
			</td>
		</tr>
		<% End If %>
        </table>
      </td>
    </tr>
    </table>
  </center>
</div>
<script language="javascript">


function onScan(ev){
var scan = ev.data;
	document.frmVars.string.value = scan.value;
	document.frmVars.submit();
}
function onSwipe(ev){
}
try
{
document.addEventListener("BarcodeScanned", onScan, false);
document.addEventListener("MagCardSwiped", onSwipe, false);
}
catch(ex) {}

<% If Request("getValType") = "" or Request("getValType") <> "" and InStr(Request("getValVal"), "*") = 0 Then %>
function ChkNum(fld, dType)
{
	if (dType != 'nvarchar')
	{
		if (!MyIsNumeric(fld.value))
		{
			alert('<%=getadSearchLngStr("DtxtValNumVal")%>');
			fld.value = '';
			fld.focus();
		}
		else if (dType == 'int')
		{
			fld.value = parseInt(fld.value);
		}
	}
}
function getValue(t, f)
{
	if (f.value != '')
	{
		document.frmVars.getValType.value = t;
		document.frmVars.getValFld.value = f.name;
		document.frmVars.getValVal.value = f.value;
		document.frmVars.cmd.value = 'adSearch';
		document.frmVars.submit();
	}
}
function doCal(varId)
{
	document.frmVars.editVar.value = varId;
	document.frmVars.cmd.value = 'adSearchValsCal';
	document.frmVars.submit();
}
function doQuery(varId)
{
	document.frmVars.editVar.value = varId;
	document.frmVars.cmd.value = 'adSearchValsQry';
	document.frmVars.submit();
}
function doChkList(varId)
{
	document.frmVars.editVar.value = varId;
	document.frmVars.cmd.value = 'adSearchValsCL';
	document.frmVars.submit();
}
function doPropList()
{
	document.frmVars.cmd.value = 'adSearchValsProp';
	document.frmVars.submit();
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
	//document.frmVars.action='adSearch.asp';
	document.frmVars.cmd.value = 'adSearch';
	document.frmVars.isSubmit.value = "R";
	document.frmVars.submit();
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
function valFrm()
{
	<% 
	doElse = False
	If rs.RecordCount > 0 Then
	rs.movefirst
	do while not rs.eof
		If rs("NotNull") = "Y" Then
			If doElse Then Response.write "else "
			If rs("Type") <> "DD" Then %>
			if (document.frmVars.var<%=rs("VarID")%>.value == '')
			{
				alert('<%=getadSearchLngStr("LtxtEnterFld")%>'.replace('{0}', '<%=rs("Name")%>'));
				<% If rs("Type") <> "CL" Then %>if (!document.frmVars.var<%=rs("VarID")%>.disabled) document.frmVars.var<%=rs("VarID")%>.focus();<% End If %>
				return false;
			}
			<% Else %>
			if (document.frmVars.var<%=rs("VarID")%>.value == '' || document.frmVars.var<%=rs("VarID")%>.value == null)
			{
				alert('|L:txtSelFld|'.replace('{0}', '<%=rs("Name")%>'));
				if (!document.frmVars.var<%=rs("VarID")%>.disabled) document.frmVars.var<%=rs("VarID")%>.focus();
				return false;
			}
		<% 	End If
			doElse = True
		End If
	rs.movenext
	loop
	End If %>
	document.frmVars.cmd.value='<%=SearchCmd%>';
	return true;
}
<% If Request("isSubmit") = "" Then %>reload('');<% End If %>
<% Else %>
function setSmallSearchVal(value)
{
	document.frmVars.<%=Request.Form("getValFld")%>.value = value;
	<% If Right(Request.Form("getValFld"), 4) = "From" Then %>
	var fldTo = document.frmVars.<%=Left(Request.Form("getValFld"), Len(Request.Form("getValFld"))-4)%>To;
	if (fldTo.value == '') fldTo.value = value;
	<% End If %>
	document.frmVars.cmd.value = 'adSearch';
	document.frmVars.submit();
}
function cancelSmallSearchVal()
{
	document.frmVars.<%=Request.Form("getValFld")%>.value = '';
	document.frmVars.cmd.value = 'adSearch';
	document.frmVars.submit();
}
<% End If %>
</script>
<% If Request("getValType") <> "" and InStr(Request("getValVal"), "*") <> 0 Then %>
<form name="frmVars" action="operaciones.asp" method="post">
<% For each itm in Request.Form
	If itm <> "getValType" and itm <> "getValFld" and itm <> "getValVal" Then %>
<input type="hidden" name="<%=itm%>" value="<%=Request.Form(itm)%>">
<% End If
Next %>
</form>
<% End If %>

<%
Function getAdSearchSQL(doQuery, Qry)
	If doQuery Then
		retVal= "declare @LanID int set @LanID = " & Session("LanID") & " declare @SlpCode int set @SlpCode = " & Session("vendid") & " " & _
		"declare @branch int set @branch = " & Session("branch") & " " & _
		"declare @CardCode nvarchar(20) set @CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "' "
	Else
		retVal = ""
	End If
	rBase.Filter = "VarID = " & rs("VarID")
	do while not rBase.eof
		If Request("var" & rBase("BaseID")) <> "" Then
			If doQuery Then
				If rBase("DataType") = "nvarchar" Then 
					MaxVar = "(" & rBase("MaxChar") & ")"
				ElseIf rBase("DataType") = "numeric" Then
					MaxVar = "(19,6)"
				Else
					MaxVar = ""
				End If
				retVal = retVal & "declare @" & rBase("Variable") & " " & rBase("DataType") & " " & MaxChar & " "
				Select Case rBase("DataType") 
					Case "nvarchar" 
						retVal = retVal & "set @" & rBase("Variable") & " = N'" & saveHTMLDecode(Request("var" & rBase("BaseID")), False) & "' "
					Case "datetime" 
						retVal = retVal & "set @" & rBase("Variable") & " = Convert(datetime,'" & SaveSqlDate(Request("var" & rBase("BaseID"))) & "',120) "
					Case Else
						retVal = retVal & "set @" & rBase("Variable") & " = " & Request("var" & rBase("BaseID")) & " "
				End Select
			Else
				retVal = retVal & "&var" & rBase("BaseID") & "=" & Request("var" & rBase("BaseID"))
			End If
		Else
			selectName = rBase("Name")
			enableControl = False
			Exit Do
		End If
	rBase.movenext
	loop
	If doQuery Then retVal = retVal & Qry
	getAdSearchSQL = retVal
End Function %>
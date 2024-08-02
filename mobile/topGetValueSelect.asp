<!--#include file="lang/topGetValueSelect.asp" -->
<%
Class clsGetValValueSelect

	Dim valType
	Dim valFld
	Dim valVal
	
	Public Property Let ValueType(ByVal p_Data)
		valType = p_Data
	End Property
	
	Public Property Let ValueField(ByVal p_Data)
		valFld = p_Data
	End Property
	
	Public Property Let Value(ByVal p_Data)
		valVal = p_Data
	End Property

	Dim onClickAction
	Dim onCancelAction
	
	Public Property Let OnClick(ByVal p_Data)
		onClickAction = p_Data
	End Property
	
	Public Property Let OnCancel(ByVal p_Data)
		onCancelAction = p_Data
	End Property

	Public Sub ShowValues
	
		set rd = Server.CreateObject("ADODB.RecordSet")
		showCol1 = False
		sql = ""
		If valType = "Crd" or valType = "CrdNam" Then
			sql = ""
			rd.close
		ElseIf valType = "ItmGrp" or valType = "ItmFrm" Then
			QryGroup = myApp.CarArt
			If Session("username") <> "" and not IsNull(Session("username")) Then
				sql = "select Case When CatalogFilterAgent = 'Y' Then CatalogFilter Else Null End CatalogFilter from OLKClientsAccess where CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "'"
				set rd = conn.execute(sql)
				If Not rd.Eof Then
					CatalogFilter = rd(0)
				End If
				rd.close
			End If
			
			sql = ""
			
			innerAddStr = ""
			If myApp.MinInvBy = "W" and myApp.GetEnableMinInv Then
				innerAddStr = "inner join OITW T____1 on T____1.ItemCode = T____0.ItemCode and T____1.WhsCode = OLKCommon.dbo.DBOLKGetWhsCode" & Session("ID") & "(" & Session("branch") & ", " & Session("vendid") & ", T____0.ItemCode) "
				fldTbl = "T____1"
			Else
				fldTbl = "T____0"
			End If
			
			If Session("PriceList") <> "" Then innerAddStr = innerAddStr & "inner join itm1 T____2 on T____2.ItemCode = T____0.ItemCode and PriceList = " & Session("PriceList") & " "
		End If
		
		searchStr = Replace(valVal,"*","%")
		Select Case valType
			Case "Crd"
				showCol1 = True
				If Not IsNull(myApp.AgentClientsFilter) Then
					AgentClientsFilter = " and CardCode collate database_default not in (" & Replace(Replace(myApp.AgentClientsFilter, "@SlpCode", Session("vendid")), "@Type", 1) & ") "
				End If
		
				sql = sql & "select CardCode, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRD', 'CardName', CardCode, CardName) CardName, " & _
							"Case CardType When 'C' Then N'" & txtClient & "' When 'S' Then N'" & gettopGetValueSelectLngStr("DtxtSupplier") & "' When 'L' Then N'" & gettopGetValueSelectLngStr("DtxtLead") & "' End [" & gettopGetValueSelectLngStr("DtxtType") & "] " & _
							"from OCRD where CardCode like N'" & searchStr & "' " & AgentClientsFilter & _
							" order by CardCode asc"
				myTitle = Replace(gettopGetValueSelectLngStr("LttlSelX"), "{0}", Server.HTMLEncode(txtClient))
				colTitle = txtClient
			Case "CrdNam"
				showCol1 = True
				If Not IsNull(myApp.AgentClientsFilter) Then
					AgentClientsFilter = " and CardCode collate database_default not in (" & Replace(Replace(myApp.AgentClientsFilter, "@SlpCode", Session("vendid")), "@Type", 1) & ") "
				End If
		
				sql = sql & "select T0.CardCode, IsNull(T1.CardName, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRD', 'CardName', T0.CardCode, T0.CardName)) CardName, " & _
							"Case T0.CardType When 'C' Then N'" & txtClient & "' When 'S' Then N'" & gettopGetValueSelectLngStr("DtxtSupplier") & "' When 'L' Then N'" & gettopGetValueSelectLngStr("DtxtLead") & "' End [" & gettopGetValueSelectLngStr("DtxtType") & "] " & _
							"from OCRD T0 " & _
							"inner join R3_ObsCommon..TDOC T1 on T1.CardCode = T0.CardCode collate database_default " & _
							"inner join R3_ObsCommon..TLOG T2 on T2.Company = db_name() and T2.LogNum = T1.LogNum and T2.Status in ('R', 'H') " & _
							"where IsNull(T1.CardName, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRD', 'CardName', T0.CardCode, T0.CardName)) like N'" & searchStr & "' " & AgentClientsFilter & _
							" Group By T0.CardCode, IsNull(T1.CardName, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRD', 'CardName', T0.CardCode, T0.CardName)), T0.CardType order by 1 asc"
				myTitle = Replace(gettopGetValueSelectLngStr("LttlSelX"), "{0}", Server.HTMLEncode(txtClient))
				colTitle = txtClient
			Case "TCrd"
				showCol1 = True
				sql = sql & "select CardCode, IsNull(CardName, '') CardName from R3_ObsCommon..TCRD T0 " & _
							"inner join R3_ObsCommon..TLOG T1 on T1.LogNum = T0.LogNum " & _
							"inner join R3_ObsCommon..TLOGControl L0 on L0.LogNum = T0.LogNum and L0.AppID = 'TM-OLK' " & _
							"where Company = db_name() and Status in ('R', 'H') and CardCode like N'" & searchStr & "%'" & _
							" order by CardCode asc"
				myTitle = Replace(gettopGetValueSelectLngStr("LttlSelX"), "{0}", Server.HTMLEncode(txtClientLead))
				colTitle = txtClientLead
			Case "TCrdNam"
				showCol1 = True
				sql = sql & "select IsNull(T0.CardCode, '') CardCode, IsNull(CardName, '') CardName " & _
							"from R3_ObsCommon..TCRD T0 " & _
							"inner join R3_ObsCommon..TLOG T2 on T2.Company = db_name() and T2.LogNum = T0.LogNum and T2.Status in ('R', 'H') " & _
							"where IsNull(CardName, '') like N'" & searchStr & "' " & _
							" Group By T0.CardCode, IsNull(CardName, '') order by 1 asc"
				myTitle = Replace(gettopGetValueSelectLngStr("LttlSelX"), "{0}", Server.HTMLEncode(txtClient))
				colTitle = txtClient
			Case "Grp"
				sql = sql & "select GroupCode, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRG', 'GroupName', GroupCode, GroupName) GroupName " & _
							"from OCRG where GroupName like N'" & searchStr & "' and GroupType = 'C' and exists(select 'A' from OCRD where GroupCode = OCRD.GroupCode) " & _
							" order by GroupName asc"
				myTitle = Replace(gettopGetValueSelectLngStr("LttlSelX"), "{0}", gettopGetValueSelectLngStr("DtxtGroup"))
				colTitle = gettopGetValueSelectLngStr("DtxtGroup")
			Case "Cty"
				showCol1 = True
				sql = sql & "select Code, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRY', 'Name', Code, Name) Name from OCRY where Name like N'" & searchStr & "'" & _
							"and exists(select 'A' from OCRD where (Country = OCRY.Code or MailCountr = OCRY.Code)) order by Name asc"
				myTitle = Replace(gettopGetValueSelectLngStr("LttlSelX"), "{0}", gettopGetValueSelectLngStr("DtxtCountry"))
				colTitle = gettopGetValueSelectLngStr("DtxtCountry")
			Case "Prj"
				showCol1 = True
				sql = sql & "select PrjCode, IsNull(PrjName, '') PrjName from OPRJ where PrjName like N'" & searchStr & "'" & _
							" order by PrjCode asc"
				myTitle = Replace(gettopGetValueSelectLngStr("LttlSelX"), "{0}", gettopGetValueSelectLngStr("LtxtPrj"))
				colTitle = gettopGetValueSelectLngStr("LtxtPrj")
			Case "Itm"
				showCol1 = True
				sql = sql & "select ItemCode, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'ItemName', ItemCode, ItemName) ItemName from OITM where ItemCode like N'" & searchStr & "%'" & _
							" order by ItemCode asc"
				myTitle = Replace(gettopGetValueSelectLngStr("LttlSelX"), "{0}", gettopGetValueSelectLngStr("DtxtItem"))
				colTitle = gettopGetValueSelectLngStr("DtxtDescription")
			Case "TItm"
				showCol1 = True
				sql = sql & "select ItemCode, IsNull(ItemName, '') ItemName from R3_ObsCommon..TITM T0 " & _
							"inner join R3_ObsCommon..TLOG T1 on T1.LogNum = T0.LogNum " & _
							"inner join R3_ObsCommon..TLOGControl L0 on L0.LogNum = T0.LogNum and L0.AppID = 'TM-OLK' " & _
							"where Company = db_name() and Status in ('R', 'H') and ItemCode like N'" & searchStr & "%' and ItemCode is not null " & _
							" order by ItemCode asc"
				myTitle = Replace(gettopGetValueSelectLngStr("LttlSelX"), "{0}", gettopGetValueSelectLngStr("DtxtItem"))
				colTitle = gettopGetValueSelectLngStr("DtxtDescription")
			Case "Slp"
				sql = sql & "select SlpCode, SlpName from OSLP where SlpName like N'" & searchStr & "%'" & _
							" order by SlpName asc"
				myTitle = Replace(gettopGetValueSelectLngStr("LttlSelX"), "{0}", Server.HTMLEncode(txtAgent))
				colTitle = txtAgent
			Case "Usr"
				sql = sql & "select INTERNAL_K, U_NAME from OUSR where U_NAME like N'" & searchStr & "%'" & _
							" order by U_NAME asc"
				myTitle = gettopGetValueSelectLngStr("LttlUser")
				colTitle = gettopGetValueSelectLngStr("LttlUser")
			Case "ItmGrp"
				itmFilter = "(select ItmsGrpCod from oitm T____0 " & innerAddStr & " " & _
							"where SellItem = 'Y' "
							
				If myApp.GetEnableMinInv Then itmFilter = itmFilter & " and " & fldTbl & ".OnHand >= " & myApp.GetMinInv & " "
							
				If Session("PriceList") <> "" Then
					itmFilter = itmFilter & "and IsNull(Price, 0)*NumInSale >= " & myApp.MinPrice & " "
				End If
						
				If QryGroup <> -1 Then	
					itmFilter = itmFilter & "and QryGroup" & QryGroup & " = 'N' "
				End If
							
				If Not IsNull(CatalogFilter) and CatalogFilter <> "" Then
					itmFilter = itmFilter & " and T____0.ItemCode not in (" & CatalogFilter & ") "
				End If
				
				If myApp.GetApplyGenFilter Then
					itmFilter = itmFilter & " and T____0.ItemCode not in (" & myApp.GetGenFilter & ") "
				End If
						
				itmFilter = itmFilter & " Group By ItmsGrpCod) "
				
				sql = sql & "select X____0.ItmsGrpCod, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITB', 'ItmsGrpNam', X____0.ItmsGrpCod, X____0.ItmsGrpNam) ItmsGrpNam from OITB X____0 "
				
				If myApp.ApplyInvFiltersBy = "I" Then
					sql = sql & " inner join " & itmFilter & " X____1 on X____1.ItmsGrpCod = X____0.ItmsGrpCod "
				End If
				
				sql = sql & " where OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITB', 'ItmsGrpNam', X____0.ItmsGrpCod, X____0.ItmsGrpNam) like N'" & searchStr & "%' "
				
				If myApp.ApplyInvFiltersBy = "N" Then
					sql = sql & " and X____0.ItmsGrpCod in " & itmFilter
				End If
							
				sql = sql & " order by ItmsGrpNam asc"
				
				myTitle = Replace(gettopGetValueSelectLngStr("LttlSelX"), "{0}", gettopGetValueSelectLngStr("LtxtItmGrp"))
				colTitle = txtAlterGrp
			Case "ItmFrm"
				itmFilter = "(select FirmCode from oitm T____0 " & innerAddStr & " " & _
							"where SellItem = 'Y' "
				
				If myApp.GetEnableMinInv Then itmFilter = itmFilter & " and " & fldTbl & ".OnHand >= " & myApp.GetMinInv & " "
							
				If Session("PriceList") <> "" Then
					itmFilter = itmFilter & "and IsNull(Price, 0)*NumInSale >= " & myApp.MinPrice & " "
				End If
						
				If QryGroup <> -1 Then	
					itmFilter = itmFilter & "and QryGroup" & QryGroup & " = 'N' "
				End If
							
				If Not IsNull(CatalogFilter) and CatalogFilter <> "" Then
					itmFilter = itmFilter & " and T____0.ItemCode not in (" & CatalogFilter & ") "
				End If
				
				If myApp.GetApplyGenFilter Then
					itmFilter = itmFilter & " and T____0.ItemCode not in (" & myApp.GetGenFilter & ") "
				End If
				
				itmFilter = itmFilter & " )"
				
				sql = sql & "select  X____0.FirmCode, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OMRC', 'FirmName',  X____0.FirmCode,  X____0.FirmName) FirmName from OMRC  X____0 "
				
				If myApp.ApplyInvFiltersBy = "I" Then
					sql = sql & " inner join " & itmFilter & " X____1 on X____1.FirmCode = X____0.FirmCode "
				End If
				
				sql = sql & " where OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OMRC', 'FirmName',  X____0.FirmCode,  X____0.FirmName) like N'" & searchStr & "%' "
				
				If myApp.ApplyInvFiltersBy = "N" Then
					sql = sql & "and  X____0.FirmCode in  " & itmFilter
				End If
							
				sql = sql & " order by X____0.FirmName asc"
				myTitle = Replace(gettopGetValueSelectLngStr("LttlSelX"), "{0}", Server.HTMLEncode(txtAlterFrm))
				colTitle = txtAlterFrm
		End Select
		set rd = conn.execute(sql) %>
		<table border="0" cellpadding="0" width="100%" id="table1">
			<tr class="CSpecialTlt">
				<td><%=myTitle%>&nbsp;</td>
			</tr>
			<tr>
				<td>
				<table border="0" width="490" id="table2" cellpadding="0">
					<% if not rd.eof then %>
					<tr class="CSpecialTlt2">
						<% If showCol1 Then %><td><%=gettopGetValueSelectLngStr("DtxtCode")%>&nbsp;</td><% End If %>
						<td><%=Server.HTMLEncode(colTitle)%>&nbsp;</td>
						<% If rd.Fields.Count >= 3 Then %><td><%=rd.Fields(2).Name%></td><% End If %>
					</tr>
					<% 
					If valType = "Grp" or valType = "Cty" or valType = "ItmGrp" or valType = "ItmFrm" or valType = "Slp" or valType = "Usr" or valType = "CrdNam" or valType = "TCrdNam" Then
						myValCol = 1
					Else
						myValCol = 0
					End If
					do while not rd.eof
					myVal = Replace(Replace(rd(myValCol), "'", "\'"), """", """""") %>
					<tr class="CSpecialTbl">
						<% If showCol1 Then %><td><a href="#" class="LinkCSpecial" onclick="<%=Replace(onClickAction, "{0}", Replace(myHTMLEncode(myVal), """", "\u0022"))%>"><%=rd(0)%></a>&nbsp;</td><% End If %>
						<td><a href="#" class="LinkCSpecial" onclick="<%=Replace(onClickAction, "{0}", Replace(myHTMLEncode(myVal), """", "\u0022"))%>"><% If Not IsNull(rd(1)) Then %><%=rd(1)%><% End If %></a>&nbsp;</td>
						<% If rd.Fields.Count >= 3 Then %><td><%=rd(rd.Fields(2).Name)%></td><% End If %>
					</tr>
					<% rd.movenext
					loop
					else %>
					<tr class="CSpecialTbl">
						<td>
						<p align="center"><%=gettopGetValueSelectLngStr("DtxtNoData")%></td>
					</tr>
					<% End If %>
				</table>
				</td>
			</tr>
			<tr class="CSpecialTbl">
				<td align="center">
				<input type="button" name="btnCancel" value="<%=gettopGetValueSelectLngStr("DtxtCancel")%>" onclick="javascript:<%=onCancelAction%>"></td>
			</tr>
		</table>
		<% 
	End Sub
End Class %>
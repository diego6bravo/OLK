<%
Class clsGetValValue 

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
	
	Public Function GetValue()
		set rVal = Server.CreateObject("ADODB.RecordSet")
		
		sql = ""
		If valType = "Crd" or valType = "CrdNam" Then
			sql = ""
			set rd = Server.CreateObject("ADODB.RecordSet")
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
			
			If Session("PriceList") <> "" Then innerAddStr = innerAddStr & "inner join itm1 T____2 on T____2.ItemCode = T____0.ItemCode and T____2.PriceList = " & Session("PriceList") & " "
		End If

		Select Case valType
			Case "DocLink"
				sql = sql & "select top 1 DocEntry, DocNum from "
				Select Case Request("DocType")
					Case 23
						sql = sql & "OQUT"
					Case 17
						sql = sql & "ORDR"
					Case 15
						sql = sql & "ODLN"
					Case 16
						sql = sql & "ORDN"
					Case 13
						sql = sql & "OINV"
					Case 14
						sql = sql & "ORIN"
					Case 203
						sql = sql & "ODPI"
					Case 24
						sql = sql & "ORCT"
					Case 46
						sql = sql & "OVPM"
					Case 67
						sql = sql & "OWTR"
				End Select
				sql = sql & " where CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "' and DocNum like '" & Request("DocNum") & "%'"
			Case "Crd"
				If Not IsNull(myApp.AgentClientsFilter) Then
					AgentClientsFilter = " and CardCode collate database_default not in (" & Replace(Replace(myApp.AgentClientsFilter, "@SlpCode", Session("vendid")), "@Type", 1) & ") "
				End If
				sql = sql & "select top 1 CardCode from OCRD where CardCode like N'" & valVal & "%'" & AgentClientsFilter & _
							" order by CardCode asc"
			Case "CrdNam"
				If Not IsNull(myApp.AgentClientsFilter) Then
					AgentClientsFilter = " and CardCode collate database_default not in (" & Replace(Replace(myApp.AgentClientsFilter, "@SlpCode", Session("vendid")), "@Type", 1) & ") "
				End If
				sql = sql & "select top 1 IsNull(T1.CardName, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRD', 'CardName', T0.CardCode, T0.CardName)) CardName " & _
							"from OCRD T0 " & _
							"inner join R3_ObsCommon..TDOC T1 on T1.CardCode = T0.CardCode collate database_default " & _
							"inner join R3_ObsCommon..TLOG T2 on T2.Company = db_name() and T2.LogNum = T1.LogNum and T2.Status in ('R', 'H') " & _
							"where IsNull(T1.CardName, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRD', 'CardName', T0.CardCode, T0.CardName)) like N'" & valVal & "%'" & AgentClientsFilter & _
							" order by 1 asc"
			Case "TCrd"
				sql = sql & "select top 1 CardCode from R3_ObsCommon..TCRD T0 " & _
							"inner join R3_ObsCommon..TLOG T1 on T1.LogNum = T0.LogNum " & _
							"inner join R3_ObsCommon..TLOGControl L0 on L0.LogNum = T0.LogNum and L0.AppID = 'TM-OLK' " & _
							"where Company = db_name() and Status in ('R', 'H') and CardCode like N'" & valVal & "%'" & _
							" order by CardCode asc"

			Case "TCrdNam"
				sql = sql & "select top 1 IsNull(CardName, '') CardName " & _
							"from R3_ObsCommon..TCRD T0 " & _
							"inner join R3_ObsCommon..TLOG T2 on T2.Company = db_name() and T2.LogNum = T0.LogNum and T2.Status in ('R', 'H') " & _
							"where IsNull(CardName, '') like N'" & valVal & "%' order by 1 asc"
			Case "Grp"
				sql = sql & "select top 1 OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRG', 'GroupName', GroupCode, GroupName) GroupName " & _
							"from OCRG where OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRG', 'GroupName', GroupCode, GroupName) like N'" & valVal & "%' and exists(select 'A' from OCRD where GroupCode = OCRD.GroupCode) " & _
							" order by GroupName asc"
			Case "Cty"
				sql = sql & "select top 1 OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRY', 'Name', Code, Name) Name " & _
							"from OCRY where OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRY', 'Name', Code, Name) like N'" & valVal & "%' and exists(select 'A' from OCRD where (Country = OCRY.Code or MailCountr = OCRY.Code)) " & _
							" order by Name asc"
			Case "Itm"
				sql = sql & "select top 1 ItemCode from OITM where ItemCode like N'" & valVal & "%'" & _
							" order by ItemCode asc"
			Case "TItm"
				sql = sql & "select top 1 ItemCode from R3_ObsCommon..TITM T0 " & _
							"inner join R3_ObsCommon..TLOG T1 on T1.LogNum = T0.LogNum " & _
							"inner join R3_ObsCommon..TLOGControl L0 on L0.LogNum = T0.LogNum and L0.AppID = 'TM-OLK' " & _
							"where Company = db_name() and Status in ('R', 'H') and ItemCode like N'" & valVal & "%'" & _
							" order by ItemCode asc"
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
						
				itmFilter = itmFilter & " Group By ItmsGrpCod)"
				
				sql = sql & "select top 1 OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITB', 'ItmsGrpNam', X____0.ItmsGrpCod, X____0.ItmsGrpNam) ItmsGrpNam from OITB X____0 "
				
				If myApp.ApplyInvFiltersBy = "I" Then
					sql = sql & " inner join " & itmFilter & " X____1 on X____1.ItmsGrpCod = X____0.ItmsGrpCod "
				End If
				
				sql = sql & " where OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITB', 'ItmsGrpNam', X____0.ItmsGrpCod, X____0.ItmsGrpNam) like N'" & valVal & "%' "
				
				If myApp.ApplyInvFiltersBy = "N" Then
					sql = sql & " and X____0.ItmsGrpCod in " & itmFilter
				End If
							
				sql = sql & " order by ItmsGrpNam asc"
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
				
				itmFilter = itmFilter & " Group By FirmCode) "
				
				sql = sql & "select top 1 OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OMRC', 'FirmName', X____0.FirmCode, X____0.FirmName)  FirmName from OMRC X____0 "
				
				If myApp.ApplyInvFiltersBy = "I" Then
					sql = sql & " inner join " & itmFilter & " X____1 on X____1.ItmsGrpCod = X____0.ItmsGrpCod "
				End If

				sql = sql & " where OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OMRC', 'FirmName', X____0.FirmCode, X____0.FirmName) like N'" & valVal & "%' "
				
				If myApp.ApplyInvFiltersBy = "N" Then
					sql = sql & " and X____0.FirmCode in " & itmFilter
				End If
				
				sql = sql & " order by FirmName asc"
			Case "Prj"
				sql = sql & "select top 1 PrjCode from OPRJ where PrjCode like N'" & valVal & "%'" & _
							" order by PrjCode asc"
			Case "AcctRejReason"
				sql = sql & "select Reason from OLKAcctRejectNotes where ReasonIndex = " & valVal
			Case "Slp"
				sql = sql & "select top 1 SlpName from OSLP where SlpName like N'" & valVal & "%'" & _
							" order by SlpName asc"
			Case "Usr"
				sql = sql & "select top 1 U_NAME from OUSR where U_NAME like N'" & valVal & "%'" & _
							" order by U_NAME asc"
		End Select
		
		set rVal = conn.execute(sql)
		
		If Not rVal.eof Then
			GetValue = rVal(0)
		Else
			GetValue = ""
		End If
		
		set rVal = Nothing
	End Function

End Class
%>
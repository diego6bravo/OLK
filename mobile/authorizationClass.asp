<%
Class clsAuthorization

 Public Property Get HasAuthorization(ByVal AutID)
 	HasAuthorization = InStr(Session("aut_" & AutID), "{Y}") <> 0 or Session("useraccess") = "P"
 End Property
 
	Public Property Get GetCardProperty(ByVal CardType, ByVal PropType)  
		Select Case CardType 
			Case "S"  
				GetCardProperty = GetObjProperty(Session("aut_78"), PropType)  
			Case "C"  
				GetCardProperty = GetObjProperty(Session("aut_45"), PropType)  
			Case "L"  
				GetCardProperty = GetObjProperty(Session("aut_77"), PropType)  
		End Select  
	End Property  

	Public Property Get GetCardType  
		tmpVal = ""
		If HasAuthorization(23) Then tmpVal = "'C'"  
		If HasAuthorization(74) Then tmpVal = ConcValue(tmpVal, "'S'") 
		If HasAuthorization(75) Then tmpVal = ConcValue(tmpVal, "'L'") 
		GetCardType = tmpVal  
	End Property  

	Public Property Get GetObjectProperty(ByVal ObjectCode, ByVal PropType) 
		Select Case ObjectCode 
			Case 4 'Item 
				GetObjectProperty = GetObjProperty(Session("aut_44"), PropType) 
			Case 13 'Invoice 
				GetObjectProperty = GetObjProperty(Session("aut_34"), PropType)
			Case -13 'Reserved Invoice 
				GetObjectProperty = GetObjProperty(Session("aut_171"), PropType) 
			Case 14 'Credit Note 
				GetObjectProperty = GetObjProperty(Session("aut_28"), PropType) 
			Case 15 'Delivery 
				GetObjectProperty = GetObjProperty(Session("aut_29"), PropType) 
			Case 16 'Return 
				GetObjectProperty = GetObjProperty(Session("aut_32"), PropType) 
			Case 17 'Sales Order 
				GetObjectProperty = GetObjProperty(Session("aut_31"), PropType) 
			Case 18 'A/P Invoice 
				GetObjectProperty = GetObjProperty(Session("aut_85"), PropType) 
			Case 19 'A/P Credit Note 
				GetObjectProperty = GetObjProperty(Session("aut_86"), PropType) 
			Case 20 'Goods Receipt PO 
				GetObjectProperty = GetObjProperty(Session("aut_83"), PropType) 
			Case 21 'Geeod Return 
				GetObjectProperty = GetObjProperty(Session("aut_84"), PropType) 
			Case 22 'Purchase Order 
				GetObjectProperty = GetObjProperty(Session("aut_82"), PropType) 
			Case 23 'Quotation 
				GetObjectProperty = GetObjProperty(Session("aut_30"), PropType) 
			Case 24 'Receipt 
				GetObjectProperty = GetObjProperty(Session("aut_27"), PropType) 
			Case 33 'Activity 
				GetObjectProperty = GetObjProperty(Session("aut_67"), PropType) 
			Case 46 'Payment 
				GetObjectProperty = GetObjProperty(Session("aut_80"), PropType) 
			Case 48 'Invoice 
				GetObjectProperty = GetObjProperty(Session("aut_35"), PropType) 
			Case 203 'A/R Down Payment Request 
				GetObjectProperty = GetObjProperty(Session("aut_176"), PropType) 
			Case 204 'A/R Down Payment Invoice 
				GetObjectProperty = GetObjProperty(Session("aut_177"), PropType) 
			Case 540000006 'Purchase Quotation 
				GetObjectProperty = GetObjProperty(Session("aut_181"), PropType) 
		End Select 
	End Property 

	Public Property Get GetObjectConfirmDocs 
		tmpVal = ""  
		If GetObjectProperty(13, "A") Then tmpVal = "13"   
		If GetObjectProperty(203, "A") Then tmpVal = ConcValue(tmpVal, "203") 
		If GetObjectProperty(204, "A") Then tmpVal = ConcValue(tmpVal, "204")
		If GetObjectProperty(15, "A") Then tmpVal = ConcValue(tmpVal, "15") 
		If GetObjectProperty(17, "A") Then tmpVal = ConcValue(tmpVal, "17") 
		If GetObjectProperty(23, "A") Then tmpVal = ConcValue(tmpVal, "23") 
		If GetObjectProperty(24, "A") Then tmpVal = ConcValue(tmpVal, "24")  
		If GetObjectProperty(22, "A") Then tmpVal = ConcValue(tmpVal, "22")  
		If GetObjectProperty(540000006, "A") Then tmpVal = ConcValue(tmpVal, "540000006") 
		GetObjectConfirmDocs = tmpVal  
	End Property 

	Public Function ConcValue(ByVal Value, ByVal AddValue) 
		If Value <> "" Then Value = Value & ", " 
		Value = Value & AddValue 
		ConcValue = Value 
	End Function 

	Private Property Get GetObjProperty(ByVal autString, ByVal PropType) 
		Select Case PropType 
			Case "C" 'Document Confirm Process 
				GetObjProperty = InStr(autString, "{C}") <> 0 and Session("useraccess") <> "P" 
			Case "A" 'Object Confirmation 
				GetObjProperty = InStr(autString, "{A}") <> 0 or Session("useraccess") = "P" 
			Case "V" 'View Object 
				GetObjProperty = InStr(autString, "{V}") <> 0 or Session("useraccess") = "P" 
			Case "S" 'Serie 
				sPos = InStr(autString, "{S}") 
             tmpVal = "" 
				If sPos Then 
					autString = Right(autString, Len(autString)-sPos-2) 
					s2Pos = InStr(autString, "{S2}") 
					If s2Pos Then autString = Left(autString, s2Pos-1) 
					tmpVal = autString 
				Else 
					tmpVal = "-1" 
				End If 
             If tmpVal = "-1" Then tmpVal = "NULL" 
             GetObjProperty = tmpVal 
			Case "S2" 'Serie 2 
				sPos = InStr(autString, "{S2}") 
             tmpVal = "" 
				If sPos Then 
					autString = Right(autString, Len(autString)-sPos-3) 
					tmpVal = autString 
				Else 
					tmpVal = "-1" 
				End If 
             If tmpVal = "-1" Then tmpVal = "NULL" 
             GetObjProperty = tmpVal 
		End Select 
	End Property 

 Public Property Get AuthorizedRepGroups
     AuthorizedRepGroups = Session("aut_repgroups")
 End Property

 Public Property Get AuthorizedForms
     AuthorizedForms = Session("aut_forms")
 End Property

 Public Property Get AuthorizedPriceList
     AuthorizedPriceList = Session("aut_pricelist")
 End Property

 Public Property Get AuthorizedBranches
     AuthorizedBranches = Session("aut_branches")
 End Property

 Public Function HasFormAuthorization(ByVal SecID) 
	returnValue = False 
 	arrForms = Split(AuthorizedForms, ", ") 
 	For i = 0 to UBound(arrForms) 
 		If CInt(arrForms(i)) = SecID Then 
 			returnValue = True 
 			Exit For 
 		End If 
 	Next 
 	HasFormAuthorization = returnValue 
 End Function 

 Public Property Get HasInAuthorization 
	returnValue = False 
	For i = 12 to 16 
		If HasAuthorization(i) Then 
			returnValue = True 
         Exit For
		End If 
	Next 
	HasInAuthorization = returnValue 
 End Property 

 Public Property Get HasOutAuthorization 
	returnValue = False 
	For i = 18 to 22 
		If HasAuthorization(i) Then 
			returnValue = True 
         Exit For
		End If 
	Next 
	HasOutAuthorization = returnValue 
 End Property 

 Public Property Get GetInOutAuthorization(ByVal By) 
 	returnValue = "" 
     Select Case By
         Case "I"
             iFrom = 12
             iTo = 16
         Case "O"
             iFrom = 18
             iTo = 22
     End Select
 	For i = iFrom to iTo 
 		If HasAuthorization(i) Then 
 			If returnValue <> "" Then returnValue = returnValue & ", " 
 			returnValue = returnValue & i 
 		End If 
 	Next 
 	GetInOutAuthorization = returnValue 
 End Property 

 Public Property Get HasBPAccess
     HasBPAccess = HasAuthorization(23) or HasAuthorization(75) or HasAuthorization(74)
 End Property

 Public Property Get HasBPCreateAccess 
     HasBPCreateAccess = HasAuthorization(77) or HasAuthorization(45) or HasAuthorization(78) 
 End Property 

	Public Property Get HasObjConfirmAccess 
		HasObjConfirmAccess = Session("useraccess") = "P" or Session("HasConfAut")
	End Property 


Sub LoadAuthorization(ByVal SlpCode, ByVal dbName) 
   	sql = "select [Authorization] from OLKAgentsAccess where SlpCode = " & SlpCode
    set rsAccess = conn.execute(sql) 

    If Not IsNull(rsAccess(0)) Then 
         arrAut = Split(rsAccess(0), "|")
         For i = 0 to UBound(arrAut) 
             autID = Split(arrAut(i), "%")(0) 
             autType = Left(autID, 1)
             endID = Replace(autID,  autType, "")

             Select Case autType
                 Case "S"
					Session("aut_" & Replace(autID, "S", "")) = arrAut(i)
                 Case "R"
                     If InStr(arrAut(i), "{Y}") Then
                         If Session("aut_repgroups") <> "" Then Session("aut_repgroups") = Session("aut_repgroups") & ", "
                         Session("aut_repgroups") = Session("aut_repgroups") & endID
                     End If 
                 Case "F"
                     If InStr(arrAut(i), "{Y}") Then
                         If Session("aut_forms") <> "" Then Session("aut_forms") = Session("aut_forms") & ", "
                         Session("aut_forms") = Session("aut_forms") & endID
                     End If 
                 Case "P"
                     If InStr(arrAut(i), "{Y}") Then
                         If Session("aut_pricelist") <> "" Then Session("aut_pricelist") = Session("aut_pricelist") & ", "
                         Session("aut_pricelist") = Session("aut_pricelist") & endID
                     End If 
                 Case "B"
                     If InStr(arrAut(i), "{Y}") Then
                         If Session("aut_branches") <> "" Then Session("aut_branches") = Session("aut_branches") & ", "
                         Session("aut_branches") = Session("aut_branches") & endID
                     End If 
             End Select
         Next 
		
		Session("HasConfAut") = InStr(Session("aut_114"), "{Y}") <> 0 or _
								InStr(Session("aut_115"), "{Y}") <> 0 or _
								InStr(Session("aut_116"), "{Y}") <> 0 or _
								InStr(Session("aut_117"), "{Y}") <> 0 or _
								InStr(Session("aut_118"), "{Y}") <> 0 or _
								InStr(Session("aut_119"), "{Y}") <> 0 or _
								InStr(Session("aut_121"), "{Y}") <> 0 or _
								InStr(Session("aut_122"), "{Y}") <> 0 or _
								InStr(Session("aut_123"), "{Y}") <> 0 or _
								InStr(Session("aut_126"), "{Y}") <> 0 or _
								InStr(Session("aut_127"), "{Y}") <> 0 or _
								InStr(Session("aut_129"), "{Y}") <> 0 or _
								InStr(Session("aut_137"), "{Y}") <> 0 or _
								InStr(Session("aut_138"), "{Y}") <> 0 or _
								InStr(Session("aut_139"), "{Y}") <> 0 or _
								InStr(Session("aut_164"), "{Y}") <> 0 or _
								InStr(Session("aut_140"), "{Y}") <> 0 or _
								InStr(Session("aut_141"), "{Y}") <> 0 or _
								InStr(Session("aut_142"), "{Y}") <> 0 or _
								InStr(Session("aut_165"), "{Y}") <> 0 or _
								InStr(Session("aut_143"), "{Y}") <> 0 or _
								InStr(Session("aut_144"), "{Y}") <> 0 or _
								InStr(Session("aut_146"), "{Y}") <> 0 or _
								InStr(Session("aut_147"), "{Y}") <> 0 or _
								InStr(Session("aut_148"), "{Y}") <> 0 or _
								InStr(Session("aut_156"), "{Y}") <> 0 or _
								InStr(Session("aut_157"), "{Y}") <> 0 or _
								InStr(Session("aut_166"), "{Y}") <> 0 or _
								InStr(Session("aut_159"), "{Y}") <> 0 or _
								InStr(Session("aut_161"), "{Y}") <> 0 or _
								InStr(Session("aut_162"), "{Y}") <> 0 or _
								InStr(Session("aut_163"), "{Y}") <> 0 or _
								InStr(Session("aut_168"), "{Y}") <> 0
								
		Session("HasActionConfAut") = 			InStr(Session("aut_115"), "{Y}") <> 0 or _
												InStr(Session("aut_117"), "{Y}") <> 0 or _
												InStr(Session("aut_119"), "{Y}") <> 0 or _
												InStr(Session("aut_122"), "{Y}") <> 0 or _
												InStr(Session("aut_123"), "{Y}") <> 0 or _
												InStr(Session("aut_127"), "{Y}") <> 0 or _
												InStr(Session("aut_129"), "{Y}") <> 0 or _
												InStr(Session("aut_138"), "{Y}") <> 0 or _
												InStr(Session("aut_139"), "{Y}") <> 0 or _
												InStr(Session("aut_164"), "{Y}") <> 0 or _
												InStr(Session("aut_141"), "{Y}") <> 0 or _
												InStr(Session("aut_142"), "{Y}") <> 0 or _
												InStr(Session("aut_165"), "{Y}") <> 0 or _
												InStr(Session("aut_144"), "{Y}") <> 0 or _
												InStr(Session("aut_146"), "{Y}") <> 0 or _
												InStr(Session("aut_156"), "{Y}") <> 0 or _
												InStr(Session("aut_157"), "{Y}") <> 0 or _
												InStr(Session("aut_166"), "{Y}") <> 0 or _
												InStr(Session("aut_159"), "{Y}") <> 0 or _
												InStr(Session("aut_161"), "{Y}") <> 0
												
		Session("HasComDocConf") = 				InStr(Session("aut_137"), "{Y}") <> 0 or _
												InStr(Session("aut_140"), "{Y}") <> 0 or _
												InStr(Session("aut_143"), "{Y}") <> 0 or _
												InStr(Session("aut_147"), "{Y}") <> 0

     End If
 End Sub


End Class
%>

<!--#include file="myHTMLEncode.asp"-->
<%
set rs = Server.CreateObject("ADODB.recordset")
AsignedSlp = not myAut.HasAuthorization(60)

Function GetTaskMonitorInfo(ByVal task)
	Select Case Task
		Case 1
			strActionAut = ""
			If Session("useraccess") = "U" Then
				If myAut.HasAuthorization(115) Then strActionAut = myAut.ConcValue(strActionAut, 115)
				If myAut.HasAuthorization(117) Then strActionAut = myAut.ConcValue(strActionAut, 117)
				If myAut.HasAuthorization(119) Then strActionAut = myAut.ConcValue(strActionAut, 119)
				If myAut.HasAuthorization(122) Then strActionAut = myAut.ConcValue(strActionAut, 122)
				If myAut.HasAuthorization(123) Then strActionAut = myAut.ConcValue(strActionAut, 123)
				If myAut.HasAuthorization(127) Then strActionAut = myAut.ConcValue(strActionAut, 127)
				If myAut.HasAuthorization(129) Then strActionAut = myAut.ConcValue(strActionAut, 129)
				If myAut.HasAuthorization(138) Then strActionAut = myAut.ConcValue(strActionAut, 138)
				If myAut.HasAuthorization(139) Then strActionAut = myAut.ConcValue(strActionAut, 139)
				If myAut.HasAuthorization(164) Then strActionAut = myAut.ConcValue(strActionAut, 164)
				If myAut.HasAuthorization(141) Then strActionAut = myAut.ConcValue(strActionAut, 141)
				If myAut.HasAuthorization(142) Then strActionAut = myAut.ConcValue(strActionAut, 142)
				If myAut.HasAuthorization(165) Then strActionAut = myAut.ConcValue(strActionAut, 165)
				If myAut.HasAuthorization(144) Then strActionAut = myAut.ConcValue(strActionAut, 144)
				If myAut.HasAuthorization(146) Then strActionAut = myAut.ConcValue(strActionAut, 146)
				If myAut.HasAuthorization(156) Then strActionAut = myAut.ConcValue(strActionAut, 156)
				If myAut.HasAuthorization(157) Then strActionAut = myAut.ConcValue(strActionAut, 157)
				If myAut.HasAuthorization(166) Then strActionAut = myAut.ConcValue(strActionAut, 166)
				If myAut.HasAuthorization(159) Then strActionAut = myAut.ConcValue(strActionAut, 159)
				If myAut.HasAuthorization(161) Then strActionAut = myAut.ConcValue(strActionAut, 161)
			End If	
	
			sql = 	"select Count('') " & _
					"from OLKUAFControl T0 " & _  
					"left outer join OSLP T1 on T1.SlpCode = T0.RequestUserSign "
					
			If Session("useraccess") = "U" Then
				If strActionAut <> "" Then
					sql = sql & "left outer join OCRD T2 on T0.ObjectCode = 2 and T2.DocEntry = T0.ObjectEntry " & _
								"inner join OLKCommon..OLKUAFActions T3 on T3.ObjCode = Case T0.ExecAt When 'O1' Then 23 When 'O0' Then 17 Else T0.ObjectCode End and (Case T0.ExecAt When 'O1' Then 23 When 'O0' Then 17 Else T0.ObjectCode End <> 2 or T0.ObjectCode = 2 and T3.CardType = T2.CardType collate database_default) and T3.ExecAt = T0.ExecAt collate database_default and T3.AutID in (" & strActionAut & ") "
				End If
			End If

			sql = sql & "left outer join R3_ObsCommon..TLOG T5 on T5.LogNum = T0.LogNum " & _
						"where T0.Status in ('O', 'P', 'E') and Left(T0.ExecAt, 1) = 'O' "
			
			If Session("useraccess") = "U" and strActionAut = "" Then
				sql = sql & " and 1 = 2 "
			End If
			
			set rs = conn.execute(sql)
			
			GetTaskMonitorInfo = CStr(rs(0))
		Case 9
			strCardType = ""
			If myAut.HasAuthorization(114) Then strCardType = myAut.ConcValue(strCardType, "'L'")
			If myAut.HasAuthorization(116) Then strCardType = myAut.ConcValue(strCardType, "'C'")
			If myAut.HasAuthorization(118) Then strCardType = myAut.ConcValue(strCardType, "'S'")
			
			sql =  "select Count('') " & _  
					"from OLKUAFControl T0    " & _  
					"inner join R3_ObsCommon..TCRD T2 on T2.LogNum = T0.ObjectEntry  " & _  
					"left outer join R3_ObsCommon..TLOG T5 on T5.LogNum = T0.ObjectEntry   " & _  
					"where ExecAt = 'C1' and T0.Status in ('O', 'P', 'E') and T5.Status = 'H' " 
					
			If Session("useraccess") = "U" Then
				sql = sql & " and T2.CardType in (" & strCardType & ") "
			End If
					
			set rs = conn.execute(sql)
			
			GetTaskMonitorInfo = CStr(rs(0))
		Case 10
			sql = 	"select Count('') " & _
					"from OLKUAFControl T0 " & _  
					"left outer join OSLP T1 on T1.SlpCode = T0.RequestUserSign " & _  
					"inner join R3_ObsCommon..TITM T2 on T2.LogNum = T0.ObjectEntry " & _  
					"left outer join OITB T3 on T3.ItmsGrpCod = T2.ItmsGrpCod " & _  
					"left outer join OMRC T4 on T4.FirmCode = T2.FirmCode " & _  
					"left outer join R3_ObsCommon..TLOG T5 on T5.LogNum = T0.ObjectEntry " & _  
					"where T0.Status in ('O', 'P', 'E') and T0.ExecAt = 'A1' and T5.Status = 'H' "
					
			set rs = conn.execute(sql)
			
			GetTaskMonitorInfo = CStr(rs(0))
		Case 11
			sql = 	"select Count('') " & _  
					"from OLKUAFControl T0  " & _  
					"left outer join OSLP T1 on T1.SlpCode = T0.RequestUserSign  " & _  
					"inner join R3_ObsCommon..TPMT T2 on T2.LogNum = T0.ObjectEntry  " & _  
					"left outer join R3_ObsCommon..TLOG T5 on T5.LogNum = T0.ObjectEntry  " & _  
					"where T0.Status in ('O', 'P', 'E') and T0.ExecAt = 'R2' and T5.Status = 'H' "
							
			set rs = conn.execute(sql)
			
			GetTaskMonitorInfo = CStr(rs(0))
		Case 12
			strObjCode = ""
			If myAut.HasAuthorization(137) Then strObjCode = myAut.ConcValue(strObjCode, "23")
			If myAut.HasAuthorization(140) Then strObjCode = myAut.ConcValue(strObjCode, "17")
			If myAut.HasAuthorization(143) Then strObjCode = myAut.ConcValue(strObjCode, "15")
			If myAut.HasAuthorization(147) Then strObjCode = myAut.ConcValue(strObjCode, "13")
			
			sql	=	"select Count('') " & _  
					"from OLKUAFControl T0  " & _  
					"left outer join OSLP T1 on T1.SlpCode = T0.RequestUserSign  " & _  
					"inner join R3_ObsCommon..TDOC T2 on T2.LogNum = T0.ObjectEntry  " & _  
					"inner join R3_ObsCommon..TLOG T6 on T6.LogNum = T0.ObjectEntry  " & _  
					"where T0.Status in ('O', 'P', 'E') and T0.ExecAt = 'D3' and T6.Status = 'H' "
			
			If Session("useraccess") = "U" Then
				sql = sql & " and T6.Object in (" & strObjCode & ") "
			End If
					
			set rs = conn.execute(sql)
			
			GetTaskMonitorInfo = CStr(rs(0))
		Case 8
			myCount = 0
				
			strGetCardType = ""
			If myAut.GetCardProperty("S", "A") Then strGetCardType = "'S'"
			If myAut.GetCardProperty("C", "A") Then strGetCardType = myAut.ConcValue(strGetCardType, "'C'")
			If myAut.GetCardProperty("L", "A") Then strGetCardType = myAut.ConcValue(strGetCardType, "'L'")
			
			If strGetCardType <> "" Then
				sql = 	"select Count('') " & _  
						"from R3_ObsCommon..TLOG T0 " & _  
						"inner join R3_ObsCommon..TLOGControl T1 on T1.LogNum = T0.LogNum and T1.AppID = 'TM-OLK' " & _  
						"inner join R3_ObsCommon..TCRD T2 on T2.LogNum = T0.LogNum " & _  
						"where T0.Status = 'H' and Company = db_name() and T0.Object = 2 and T2.CardType in (" & strGetCardType & ") " 
				set rs = conn.execute(sql)
				myCount = myCount + CInt(rs(0))
			End If
				
					
			strObjCode = ""
			If myAut.GetObjectProperty(4, "A") Then strObjCode = "4"
			If myAut.GetObjectProperty(13, "A") Then strObjCode = myAut.ConcValue(strObjCode, "13")
			If myAut.GetObjectProperty(15, "A") Then strObjCode = myAut.ConcValue(strObjCode, "15")
			If myAut.GetObjectProperty(17, "A") Then strObjCode = myAut.ConcValue(strObjCode, "17")
			If myAut.GetObjectProperty(23, "A") Then strObjCode = myAut.ConcValue(strObjCode, "23")
			If myAut.GetObjectProperty(24, "A") Then strObjCode = myAut.ConcValue(strObjCode, "24")
			
			If Task = 8 Then slpFilter = " and T1.SlpCode = " & Session("vendid") & " "
					
			If strObjCode <> "" Then
				sql = 	"select Count('') " & _  
						"from r3_obscommon..tlog T0   " & _  
						"inner join R3_ObsCommon..TLOGControl T1 on T1.LogNum = T0.LogNum and T1.AppID = 'TM-OLK' " & slpFilter & _  
						"where Company = db_name() and Status = 'H' and Object in (" & strObjCode & ")  " 
				set rs = conn.execute(sql)
				myCount = myCount + CInt(rs(0))
			End If
			
			GetTaskMonitorInfo = CStr(myCount)
		Case 2
			sql = "select Count('') from OCRD T0 inner join OLKClientsAccess T3 on T3.CardCode = T0.CardCode where T3.Status = 'P' "
			set rs = conn.execute(sql)
			
			GetTaskMonitorInfo = CStr(rs(0))
		Case 3
			sql = 	"declare @SlpCode int set @SlpCode = " & Session("vendid") & " " & _  
					"select Sum([Count]) " & _  
					"from " & _  
					"( " & _  
					"	select Count('') [Count] " & _  
					"	from R3_Obscommon..tlog T0  " & _  
					"	inner join r3_obscommon..TCLG T1 on T1.LogNum = T0.LogNum  " & _  
					"	inner join R3_ObsCommon..TLOGControl X0 on X0.LogNum = T0.LogNum and X0.appId = 'TM-OLK'  " & _  
					"	where Company = db_name() and Object = 33 and T0.status = 'R' and T1.SlpCode = @SlpCode " & _  
					"	union all " & _  
					"	select Count('') " & _  
					"	from OCLG T1  " & _  
					"	inner join ocrd T2 on T2.CardCode = T1.CardCode collate database_default  " & _  
					"	left outer join ocry T3 on T3.code = T2.Country collate database_default  " & _  
					"	inner join ocrg T4 on T4.GroupCode = T2.GroupCode  " & _  
					"	left outer join OUSR T6 on T6.INTERNAL_K = T1.AttendUser  " & _  
					"	left outer join OHEM T7 on T7.userId = T1.AttendUser  " & _  
					"	where T1.Closed = 'N' and T1.Inactive = 'N' and T7.salesPrson = @SlpCode " & _  
					") X0 "
			set rs = conn.execute(sql)
			GetTaskMonitorInfo = CStr(rs(0))
		Case 4
			sql = 	"declare @SlpCode int set @SlpCode = " & Session("vendid") & " " & _  
					"select top 1 [Source], TransNum, CardCode " & _  
					"from ( " & _  
					"	select 'O' [Source], T0.LogNum TransNum, T1.Recontact,  " & _  
					"	Replace(Left(Convert(nvarchar(10),BeginTime,108), Len(Convert(nvarchar(10),BeginTime,108)) - 3), ':', '') [BeginTime], T1.CardCode collate database_default CardCode " & _  
					"	from R3_Obscommon..tlog T0  " & _  
					"	inner join r3_obscommon..TCLG T1 on T1.LogNum = T0.LogNum  " & _  
					"	inner join ocrd T2 on T2.CardCode = T1.CardCode collate database_default  " & _  
					"	left outer join ocry T3 on T3.code = T2.Country collate database_default  " & _  
					"	inner join ocrg T4 on T4.GroupCode = T2.GroupCode  " & _  
					"	left outer join oslp T5 on T5.slpcode = T1.SlpCode  " & _  
					"	left outer join OUSR T6 on T6.INTERNAL_K = T1.AttendUser  " & _  
					"	inner join R3_ObsCommon..TLOGControl X0 on X0.LogNum = T0.LogNum and X0.appId = 'TM-OLK'  " & _  
				"	where Company = db_name() and Object = 33 and T0.status = 'R' and T1.SlpCode = @SlpCode and DateDiff(day, getdate(),T1.Recontact) >= 0 " & _  
				"	union all " & _  
				"	select 'S' [Souce], T1.ClgCode TransNum, T1.Recontact, BeginTime, T1.CardCode " & _  
				"	from OCLG T1  " & _  
				"	inner join ocrd T2 on T2.CardCode = T1.CardCode collate database_default  " & _  
				"	left outer join ocry T3 on T3.code = T2.Country collate database_default " & _  
				"	 inner join ocrg T4 on T4.GroupCode = T2.GroupCode  " & _  
				"	left outer join OUSR T6 on T6.INTERNAL_K = T1.AttendUser  " & _  
				"	left outer join OHEM T7 on T7.userId = T1.AttendUser  " & _  
				"	where T1.Closed = 'N' and T1.Inactive = 'N' and T7.salesPrson = @SlpCode and DateDiff(day, getdate(),T1.Recontact) >= 0 " & _  
				") X0 " & _  
				"where DateDiff(day,getdate(),Recontact) > 0 or " & _  
				"DateDiff(day,Recontact,getdate()) = 0 and BeginTime >= Replace(Left(Convert(nvarchar(10),getdate(),108), Len(Convert(nvarchar(10),getdate(),108)) - 3), ':', '') " & _  
				"order by Recontact, BeginTime "
			set rs = conn.execute(sql)
			Dim retVal(3)
			If Not rs.Eof Then
				retVal(0) = rs(0)
				retVal(1) = rs(1)
				retVal(2) = rs(2)
			Else
				retVal(1) = -1
			End If
			GetTaskMonitorInfo  = retVal
		Case 5

			ObjCode = ""
			If myApp.EnableOQUT Then ObjCode = "23"
			If myApp.EnableORDR Then ObjCode = myAut.ConcValue(ObjCode, "17")
			If myApp.EnableODLN Then ObjCode = myAut.ConcValue(ObjCode, "15")
			If myApp.EnableODPIReq Then ObjCode = myAut.ConcValue(ObjCode, "203")
			If myApp.EnableODPIInv Then ObjCode = myAut.ConcValue(ObjCode, "204")
			If myApp.EnableOINV or myApp.EnableCashInv or myApp.EnableOINVRes Then ObjCode = myAut.ConcValue(ObjCode, "13")
			If myApp.EnableORCT Then ObjCode = myAut.ConcValue(ObjCode, "24")
		
			If ObjCode <> "" Then
				set cmd = Server.CreateObject("ADODB.Command")
				cmd.ActiveConnection = connCommon
				cmd.CommandType = &H0004
				cmd.CommandText = "DBOLKGetTaskMonitorDocCount" & Session("ID")
				cmd.Parameters.Refresh()
				cmd("@AllowAgentAccessCDoc") = GetYN(myApp.AllowAgentAccessCDoc)
				cmd("@SlpCode") = Session("vendid")
				cmd("@Objects") = ObjCode
				cmd("@EnableInv") = GetYN(myApp.EnableOINV)
				cmd("@EnableCashInv") = GetYN(myApp.EnableCashInv)
				cmd("@EnableInvRes") = GetYN(myApp.EnableOINVRes)

				set rs = cmd.execute()
				
				GetTaskMonitorInfo = CInt(rs(0))
				
			Else
				GetTaskMonitorInfo = -1
			End If
		Case 6

				sql = 	"select T0.AdPollID, T0.Filter " & _
						"from OLKADPoll T0 " & _
						"inner join OLKADPollAgents T1 on T1.AdPollID = T0.AdPollID and (T1.SlpCode = " & Session("vendid") & " or T1.SlpCode = -2) " & _
						"where DateDiff(day,StartDate,getdate()) >= 0 and DateDiff(day,getdate(),EndDate) >= 0 and Status = 'A'"
				set rs = conn.execute(sql) 
			
				If Not rs.Eof Then
					sql = 	"select Count('') [Count], Sum(Pending) Pending from (" & _
							"SELECT T0.ADPollID, Case T0.ADPollID "
					
			
					do while not rs.eof
						If rs("Filter") <> "" Then qcFilter = " and " & rs("Filter") else qcFilter = ""
						sql = sql & "When " & rs("ADPollID") & " Then (select count('') from OCRD where CardType in ('C', 'L') " & qcFilter & " and not exists(select 'S' from OLKADPollAnswers where ADPollID = T0.ADPollID and CardCode = OCRD.CardCode)) "	
					rs.movenext
					loop
						
					sql = sql & " End Pending from OLKADPoll T0 " & _
					"left outer join OLKADPollAlterNames T1 on T1.LanID = " & Session("LanID") & " and T1.AdPollID =  T0.AdPollID " & _
					"where DateDiff(day,StartDate,getdate()) >= 0 and DateDiff(day,getdate(),EndDate) >= 0 and Status = 'A' " & _
					"Group By T0.ADPollID) X0"
			
					set rs = conn.execute(sql)
					
					GetTaskMonitorInfo = rs(0) & "&nbsp;(" & rs(1) & ")"
				Else
					GetTaskMonitorInfo = "-1"
				End If
		Case 7
				sql = 	"select IsNull(Sum(Case When T0.ofertStatus in ('O', 'W') Then 1 Else 0 End), 0) TotalWait, Count('') Total " & _  
						"from OLKOferts T0  " & _  
						"inner join OLKOfertsLines T1 on T1.OfertIndex = T0.OfertIndex  " & _  
						"inner join OCRD T2 on T2.CardCode = T0.UserName  " & _  
						"where OfertLineNum = (select max(OfertLineNum) from OLKOfertsLines where OfertIndex = T0.OfertIndex) and TransStatus = 'O' and T0.OfertStatus not in ('A', 'R') " 
						
				If Not IsNull(myApp.AgentClientsFilter) and not IgnoreGeneralFilter Then
					sql = sql & " and T0.UserName not in (" & Replace(Replace(myApp.AgentClientsFilter, "@SlpCode", Session("vendid")), "@Type", 1) & ") "
				End If
				
				If not myAut.HasAuthorization(60) Then sql = sql & " and T2.SLPCode = " & Session("vendid") & " "
				
				set rs = conn.execute(sql)
				
				Dim retValOffers(2)
				retValOffers(0) = rs(0)
				retValOffers(1) = rs(1)
				
				GetTaskMonitorInfo = retValOffers
		Case 13
				sql = "declare @SlpCode int set @SlpCode = " & Session("vendid") & " " & _  
						"select " & _  
						"(select Count('') " & _  
						"from R3_Obscommon..tlog T0 " & _  
						"inner join r3_obscommon..TOPR T1 on T1.LogNum = T0.LogNum " & _  
						"left outer join OHEM T6 on T6.empID = T1.Owner " & _  
						"inner join R3_ObsCommon..TLOGControl X0 on X0.LogNum = T0.LogNum and X0.appId = 'TM-OLK' " & _  
						"where Company = N'" & Session("olkdb") & "' and Object = 97 and T0.status = 'R' and T1.SlpCode = @SlpCode) + " & _  
						"(select Count('') from OOPR T1 where T1.Status = 'O'  and T1.SlpCode = @SlpCode) " 
			set rs = conn.execute(sql)
			GetTaskMonitorInfo = CStr(rs(0))
		Case 14, 15, 16, 17, 18
			ExecAt = ""
			Select Case Task
				Case 14
					ExecAt = "O"
				Case 15
					ExecAt = "C1"
				Case 16
					ExecAt = "A1"
				Case 17
					ExecAt = "R2"
				Case 18
					ExecAt = "D3"
			End Select 
			
			sql = "declare @SlpCode int set @SlpCode = " & Session("vendid") & " " & _  
					"select Count('') [Count] " & _  
					"from OLKUAFControl T0   " & _  
					"left outer join R3_ObsCommon..TLOG T5 on T5.LogNum = T0.ObjectEntry " & _  
					"inner join OLKUAFControl2 X0 on X0.ID = T0.ID " & _  
					"inner join OLKUAF4 X1 on X1.FlowID = X0.FlowID and X1.AutGrpID = X0.AutGrpID " & _  
					"inner join OLKAutGrpSlp X2 on X2.GrpID = X0.AutGrpID and X2.SlpCode = @SlpCode " & _  
					"inner join OLKUAF X3 on X3.FlowID = X0.FlowID " & _  
					"where T0.Status = 'O' and (T5.Status = 'H' or LEFT(T0.ExecAt, 1) = 'O') and X0.Status = 'W' " & _  
					"and  " & _  
					"( " & _  
					"	(select Status from OLKUAFControl2 where ID = X0.ID and FlowID = X0.FlowID and LineID = X0.LineID-1) is null " & _  
					"	or " & _  
					"	(select Status from OLKUAFControl2 where ID = X0.ID and FlowID = X0.FlowID and LineID = X0.LineID-1) = 'A' " & _  
					") " & _  
					"and Case When Left(T0.ExecAt, 1) = 'O' Then 'O' Else T0.ExecAt End = '" & ExecAt & "' " 
			set rs = Server.CreateObject("ADODB.RecordSet")
			set rs = conn.execute(sql)
			GetTaskMonitorInfo = CStr(rs(0))
	End Select
End Function
%>
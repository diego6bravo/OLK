<% addLngPathStr = "payments/" %>
<!--#include file="lang/finalPayment.asp" -->
<% 

			sql = "select T0.TableID, T0.FieldID, AliasID, (select SDKID Collate database_default from r3_obscommon..tcif where companydb = N'" & Session("OLKDb") & "')++AliasID As InsertID, " & _
				  "case when exists(select 'A' from UFD1 where TableID = T0.TableID and FieldID = T0.FIeldID) Then 'Y' Else 'N' End UFD1, TypeID, EditType, RTable " & _
				  "from cufd T0 " & _
				  "left outer join OLKCUFD T1 on T1.TableID = T0.TableID and T1.FieldID = T0.FieldID " & _
				  "where T0.TableId = 'ORCT' and AType in ('V','T') and OP in ('T','O') and Active = 'Y'"
			set rs = server.createobject("ADODB.RecordSet")
			rs.open sql, conn, 3, 1
			
			sqlFinalRecAdd  = ""
			do while not rs.eof
				If rs("UFD1") = "N" and IsNull(rs("RTable")) Then
					sqlFinalRecAdd = sqlFinalRecAdd & ", " & rs("InsertID") & " U_" & rs("AliasID")
				ElseIf not IsNull(rs("RTable")) Then
					sqlFinalRecAdd = sqlFinalRecAdd & ", (select Name from [@" & rs("RTable") & "] where Code = T0." & rs("InsertID") & " collate database_default) U_" & rs("AliasID") & ""
				Else
					sqlFinalRecAdd = sqlFinalRecAdd & ", (select IsNull((select AlterDescr from OLKUFD1AlterNames where TableID = UFD1.TableID and FieldID = UFD1.FieldID and IndexID = UFD1.IndexID and LanID = " & Session("LanID") & "),Descr) " & _
														" from UFD1 where TableID = '" & rs("TableID") & "' and FieldID = " & rs("FieldID") & " and FldValue collate database_default = T0." & rs("InsertID") & ") U_" & rs("AliasID") & ""
				End If
			rs.movenext
			loop
			rs.close
			
           set rc = Server.CreateObject("ADODB.recordset")
           sql = "EXEC OLKCommon..DBOLKGetPymntConfirmInfo" & Session("ID") & " " & Session("ConfPayRetVal") & ", N'" & Replace(sqlFinalRecAdd,"'","''") & "'"
           response.write sql
           set rc = conn.execute(Sql)
           EnableSDK = rc("EnableSDK")

	       If rc("DocType") = 13 then
		       DocType = 13
		       docName = txtInv '"Factura"
	       ElseIf rc("DocType") = 17 Then
		       DocType = 17
		       docName = txtOrdr '"Pedido"
	       End If
%>
<table border="0" cellpadding="0" width="100%">
	<% If rc("PrintCmpPaper") = "Y" Then %>
	<tr class="GeneralTltBig">
		<td height="80">&nbsp;</td>
	</tr>
	<% If rc("Draft") = "Y" Then %>
	<tr>
		<td class="TablasTituloDraft">
		<%=getfinalPaymentLngStr("LttlDraftNote")%></td>
	</tr>
	<% End If %>
	<% Else
	set ra = Server.CreateObject("ADODB.RecordSet")
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKGetObjectData" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@ObjType") = "S"
	cmd("@ObjID") = 19
	cmd("@UserType") = userType
	set ra = cmd.execute()
	strContent = ra("ObjContent")
	strContent = Replace(strContent, "{SelDes}", SelDes)
	strContent = Replace(strContent, "{rtl}", Session("rtl"))  
	strContent = Replace(strContent, "{CmpName}", CmpName) %>
	<%=strContent%>
	<!--#include file="../myTitle.asp"-->
	<tr class="TablasTituloSec">
		<td>
		<% If Request("Confirm") <> "H" Then %><%=txtRct%><% If rc("Draft") = "Y" Then %> (<%=getfinalPaymentLngStr("DtxtDraft")%>)<% End If %> #<%=rc("DocNum")%><% Else %><%=getfinalPaymentLngStr("DtxtLogNum")%><%=Session("ConfPayRetVal")%><% End If %></td>
	</tr>
	<% If rc("Draft") = "Y" Then %>
	<tr>
		<td class="TablasTituloDraft">
		<%=getfinalPaymentLngStr("LttlDraftNote")%></td>
	</tr>
	<% End If %>
	<% End If %>
</table>
<table border="0" cellpadding="0" width="100%">
	<tr>
		<td width="50%" valign="top">
		<table border="0" cellpadding="0" width="100%" cellspacing="1" id="table4">
			<tr>
				<td class="GeneralTblBold2" width="100"><%=getfinalPaymentLngStr("DtxtCode")%></td>
				<td class="FirmTbl"><%=rc("CardCode")%>&nbsp;</td>
			</tr>
			<tr>
				<td class="GeneralTblBold2" width="100"><%=getfinalPaymentLngStr("DtxtName")%></td>
				<td class="FirmTbl"><%=rc("CardName")%>&nbsp;</td>
			</tr>
			<% if rc("PrintAddress") = "Y" then%>
			<tr>
				<td class="GeneralTblBold2" width="100"><%=getfinalPaymentLngStr("DtxtAddress")%></td>
				<td class="FirmTbl"><%=rc("Address")%>&nbsp;</td>
			</tr>
			<% end if %>
			<% if rc("PrintContact") = "Y" then %>
			<tr>
				<td class="GeneralTblBold2" width="100"><%=getfinalPaymentLngStr("DtxtContact")%></td>
				<td class="FirmTbl"><% If Not IsNull(rc("CntctName")) Then %><%=rc("CntctName")%><% End If %>&nbsp;</td>
			</tr>
			<%end if %>

            <% If EnableSDK = "Y" Then
	sql = "select T0.FieldID, AliasID, IsNull(AlterDescr, Descr) Descr, TypeID, SizeID, Dflt, NotNull, Pos, EditType, " & _
		  "Case When Exists(select 'A' from UFD1 where TableId = T0.TableId and FieldId = T0.FieldId) " & _
		  "Then 'Y' Else 'N' End As DropDown, NullField, Query, " & _
		  "(select SDKID collate database_default from r3_obscommon..tcif where companydb = N'" & Session("OlkDB") & "')++AliasID As InsertID " & _
		  "from cufd T0 " & _
		  "left outer join OLKCUFD T1 on T1.TableID = T0.TableID and T1.FieldID = T0.FieldID " & _
		  "left outer join OLKCUFDAlterNames T2 on T2.TableID = T0.TableID and T2.FieldID = T0.FieldID and T2.LanID = " & Session("LanID") & " " & _
		  "left outer join OLKCUFDGroups T3 on T3.TableID = T1.TableID and T3.GroupID = IsNull(T1.GroupID, -1) " & _
		  "where T0.TableId = 'ORCT' and AType in ('V', 'T') and OP in ('T','O') and Active = 'Y' " & _
		  "order by T3.[Order], IsNull(T1.Pos, 'D'), IsNull(T1.[Order], 32727) "
			set rs = Server.CreateObject("ADODB.RecordSet")
			rs.open sql, conn, 3, 1
			rs.Filter = "Pos = 'I'"
			do while not rs.eof %>
            <tr>
              <td width="25%" class="GeneralTblBold2">
            <%=rs("Descr")%>&nbsp;</td>
              <td width="75%" class="FirmTbl">
              <% If rs("TypeID") = "M" and rs("EditType") = "B" Then %><a target="_blank" href="<%=rc("U_" & rs("AliasID"))%>"><% End If %>
              <% If rs("TypeID") = "B" Then
              		If Not IsNull(rc("U_" & rs("AliasID"))) Then
	            	Select Case rs("EditType")
						Case "R"
							Response.Write FormatNumber(CDbl(rc("U_" & rs("AliasID"))),myApp.RateDec)
						Case "S"
							Response.Write FormatNumber(CDbl(rc("U_" & rs("AliasID"))),myApp.SumDec)
						Case "P"
							Response.Write FormatNumber(CDbl(rc("U_" & rs("AliasID"))),myApp.PriceDec)
						Case "Q"
							Response.Write FormatNumber(CDbl(rc("U_" & rs("AliasID"))),myApp.QtyDec)
						Case "%"
							Response.Write FormatNumber(CDbl(rc("U_" & rs("AliasID"))),myApp.PercentDec)
						Case "M"
							Response.Write FormatNumber(CDbl(rc("U_" & rs("AliasID"))),myApp.MeasureDec)
	            	End Select
	            	End If
              ElseIf rs("TypeID") = "A" and rs("EditType") = "I" Then %>
              <img src="../pic.aspx?filename=<% If Not IsNull(rc("U_" & rs("AliasID"))) Then %><%=rc("U_" & rs("AliasID"))%><% Else %>n_a.gif<% End If %>&MaxSize=180&dbName=<%=Session("olkdb")%>" border="0">
              <% ElseIf rs("TypeID") = "D" Then %>
              <%=FormatDate(rc("U_" & rs("AliasID")), True)%>
              <% Else %>
              <% If Not IsNull(rc("U_" & rs("AliasID"))) Then %><%=rc("U_" & rs("AliasID"))%><% End If %>
              <% End If %>
              <% If rs("TypeID") = "M" and rs("EditType") = "B" Then %></a><% End If %>
              </td>
            </tr>
            <% rs.movenext
            loop
            End If %>
			<% If rc("DocNum") <> "" and not IsNull(rc("DocNum")) Then %>
			<tr>
				<td class="GeneralTblBold2" colspan="2"><% If 1 = 2 Then %>Recibo<% Else %><%=txtRct%><% End If %> 
				#<%=rc("DocNum")%>&nbsp;</td>
			</tr>
			<% End If %>
		</table>
		</td>
		<td width="50%" valign="top">
		<table border="0" cellpadding="0" width="100%" cellspacing="1" id="table3">
			<tr>
				<td class="GeneralTblBold2" width="130"><%=getfinalPaymentLngStr("DtxtDate")%></td>
				<td class="FirmTbl"><%=FormatDate(rc("DocDate"), True)%>&nbsp;</td>
			</tr>
			<tr>
				<td class="GeneralTblBold2" width="130"><%=getfinalPaymentLngStr("LtxtCounterRef")%></td>
				<td class="FirmTbl"><%=rc("CounterRef")%>&nbsp;</td>
			</tr>
			<% If EnableSDK = "Y" Then
				rs.Filter = "Pos = 'D'"
				if rs.recordcount > 0 then rs.movefirst
				do while not rs.eof %>
	            <tr>
	              <td width="25%" class="GeneralTblBold2">
        		<%=rs("Descr")%>&nbsp;</td>
	              <td width="75%" class="FirmTbl">
	              <% If rs("TypeID") = "M" and rs("EditType") = "B" Then %><a target="_blank" href="<%=rc("U_" & rs("AliasID"))%>"><% End If %>
                  <% If rs("TypeID") = "B" Then
                  		If Not IsNull(rc("U_" & rs("AliasID"))) Then
		            	Select Case rs("EditType")
							Case "R"
								Response.Write FormatNumber(CDbl(rc("U_" & rs("AliasID"))),myApp.RateDec)
							Case "S"
								Response.Write FormatNumber(CDbl(rc("U_" & rs("AliasID"))),myApp.SumDec)
							Case "P"
								Response.Write FormatNumber(CDbl(rc("U_" & rs("AliasID"))),myApp.SumDec)
							Case "Q"
								Response.Write FormatNumber(CDbl(rc("U_" & rs("AliasID"))),myApp.QtyDec)
							Case "%"
								Response.Write FormatNumber(CDbl(rc("U_" & rs("AliasID"))),myApp.PercentDec)
							Case "M"
								Response.Write FormatNumber(CDbl(rc("U_" & rs("AliasID"))),myApp.MeasureDec)
		            	End Select
		            	End If
                  ElseIf rs("TypeID") = "A" and rs("EditType") = "I" Then %>
                  <img src="../pic.aspx?filename=<% If Not IsNull(rc("U_" & rs("AliasID"))) Then %><%=rc("U_" & rs("AliasID"))%><% Else %>n_a.gif<% End If %>&MaxSize=180&dbName=<%=Session("olkdb")%>" border="0">
	             <% ElseIf rs("TypeID") = "D" Then %>
	              <%=FormatDate(rc("U_" & rs("AliasID")), True)%>
                  <% Else %>
                  <% If Not IsNull(rc("U_" & rs("AliasID"))) Then %><%=rc("U_" & rs("AliasID"))%><% End If %>
                  <% End If %>
	              <% If rs("TypeID") = "M" and rs("EditType") = "B" Then %></a><% End If %>
	              </td>
	            </tr>
	            <% rs.movenext
	            loop
	            End If %>
			</table>
		</td>
	</tr>
	<tr>
		<td width="99%" colspan="2">
		<table border="0" cellpadding="0" width="100%" id="table5">
			<tr class="GeneralTlt">
				<td align="center" style="width: 20%"><%=getfinalPaymentLngStr("DtxtDoc")%></td>
				<td align="center" style="width: 20%"><%=getfinalPaymentLngStr("DtxtInstallment")%></td>
				<td align="center"><%=getfinalPaymentLngStr("DtxtDate")%></td>
				<td align="center"><%=getfinalPaymentLngStr("DtxtDetail")%></td>
				<td align="center"><%=getfinalPaymentLngStr("DtxtTotal")%></td>
				<td align="center" width="124"><%=getfinalPaymentLngStr("DtxtPaid")%></td>
				<td align="center" width="120"><%=getfinalPaymentLngStr("DtxtBalance")%></td>
			</tr>
      <% 
       set rs = Server.CreateObject("ADODB.recordset")
       sql = "declare @MainCur nvarchar(3) set @MainCur = (select top 1 MainCurncy from oadm) "

       sql = sql & "select DocNum, T2.InstlmntID InstID, T2.DueDate DocDate, IsNull(T1.Comments, '') Comments, " & _
				"Case When T1.DocCur = @MainCur Then T2.InsTotal Else T2.InsTotalFC End DocTotal, SumApplied, " & _
				"Case When T1.DocCur = @MainCur Then T2.InsTotal Else T2.InsTotalFC End -  " & _
				"Case When T1.DocCur = @MainCur Then T2.PaidToDate Else T2.PaidFC End Saldo, T1.DocCur, 1 InstCount  " & _
				"from R3_Obscommon..pmt2 T0  " & _
				"inner join oinv T1 on T1.DocEntry = T0.DocEntry " & _
				"inner join inv6 T2 on T2.DocEntry = T0.DocEntry and T2.InstlmntID = T0.InstID " & _
				"where T0.lognum = " & Session("ConfPayRetVal")
			rs.open sql, conn, 3, 1
			TotalSaldo = 0
			do while not rs.eof 
			TotalSaldo = TotalSaldo + CDbl(rs("SumApplied"))
			TotalCancelado = TotalCancelado + CDbl(rs("Saldo")) %>
			<tr class="GeneralTbl">
				<td style="width: 20%"><%=rs("DocNum")%>&nbsp;</td>
				<td style="width: 20%"><%=Replace(Replace(getfinalPaymentLngStr("DtxtXofY"), "{0}", rs("InstID")), "{1}", rs("InstCount"))%>&nbsp;</td>
				<td><%=FormatDate(rs("DocDate"), True)%>&nbsp;</td>
				<td><%=rs("Comments")%>&nbsp;</td>
				<td align="right"><nobr><%=rs("DocCur")%>&nbsp;<%=FormatNumber(rs("DocTotal"),myApp.SumDec)%></nobr></td>
				<td width="124" align="right"><nobr><%=rs("DocCur")%>&nbsp;<%=FormatNumber(rs("SumApplied"),myApp.SumDec)%></nobr></td>
				<td align="right" width="120"><nobr><%=rs("DocCur")%>&nbsp;<%=FormatNumber(rs("Saldo"),myApp.SumDec)%></nobr></td>
			</tr>
			<% rs.movenext
			loop %>
			<tr>
				<td colspan="5" rowspan="3" valign="top">
				<table border="0" cellpadding="0" width="75%" cellspacing="1" id="table6">
					<% if CDbl(Rc("CashSum")) <> 0 then %>
					<tr class="GeneralTbl">
						<td class="GeneralTblBold2"  width="103"><%=getfinalPaymentLngStr("LtxtCash")%> </td>
						<td>&nbsp;</td>
						<td>&nbsp;</td>
						<td align="right"><nobr><%=rc("DocCur")%>&nbsp;<%=FormatNumber(CDbl(Rc("CashSum")),myApp.SumDec)%></nobr></td>
					</tr>
					<% End If %>
					<% if CDbl(Rc("CheckSum")) <> 0 then
					sql = "select  IsNull((select BankName from odsc where BankCode = T0.BankCode collate database_default), '') BankName, CheckNum,CheckSum " & _
					"from R3_ObsCommon..PMT1 T0 " & _
					"where LogNum = " & Session("ConfPayRetVal") 
					set rs = conn.execute(sql)
					do while not rs.eof %>
					<tr class="GeneralTbl">
						<td class="GeneralTblBold2"  width="103"><%=getfinalPaymentLngStr("LtxtCheck")%></td>
						<td><%=rs("BankName")%>&nbsp;</td>
						<td>#<%=rs("CheckNum")%></td>
						<td align="right"><nobr><%=rc("DocCur")%>&nbsp;<%=FormatNumber(CDbl(Rs("CheckSum")),myApp.SumDec)%></nobr></td>
					</tr>
					<% rs.movenext
					loop 
					end if %>
					<% if CDBl(Rc("CreditSum")) <> 0 then
					sql = "select IsNull((select OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRC', 'CardName', T0.CreditCard, CardName) from OCRC where CreditCard = T0.CreditCard), '') CardName, CreditSum " & _
					"from R3_ObsCommon..PMT3 T0 " & _
					"where LogNum = " & Session("ConfPayRetVal")
					set rs = conn.execute(sql)
					do while not rs.eof %>
					<tr class="GeneralTbl">
						<td class="GeneralTblBold2"  width="103"><%=getfinalPaymentLngStr("LtxtCredCard")%></td>
						<td><%=rs("CardName")%>&nbsp;</td>
						<td>&nbsp;</td>
						<td align="right"><nobr><%=rc("DocCur")%>&nbsp;<%=FormatNumber(CDbl(Rc("CreditSum")),myApp.SumDec)%></nobr></td>
					</tr>
					<% rs.movenext
					loop
					end if %>
					<% if CDbl(Rc("TrsfrSum")) <> 0 then%>
					<tr class="GeneralTbl">
						<td class="GeneralTblBold2"  width="103">
						<%=getfinalPaymentLngStr("LtxtBankTrans")%></td>
						<td>&nbsp;</td>
						<td>&nbsp;</td>
						<td>
						<p align="right"><nobr><%=rc("DocCur")%>&nbsp;<%=FormatNumber(CDbl(Rc("TrsfrSum")),myApp.SumDec)%></nobr></td>
					</tr>
					<% end if %>
				</table>
				</td>
				<td class="GeneralTblBold2"  width="124"><% If rc("PrintPaidSum") = "N" Then %><%=getfinalPaymentLngStr("LtxtPrevBal")%><% End If %></td>
				<td class="GeneralTbl" align="right" width="120" ><% If rc("PrintPaidSum") = "N" Then %><nobr><%=rc("DocCur")%>&nbsp;<%=FormatNumber(TotalSaldo+TotalCancelado,myApp.SumDec)%></nobr><% End If %></td>
			</tr>
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2" width="124"><%=getfinalPaymentLngStr("DtxtImport")%></td>
				<td class="GeneralTbl" width="120" align="right"><nobr><%=rc("DocCur")%>&nbsp;<%=FormatNumber(rc("Pagado"),myApp.SumDec)%></nobr></td>
			</tr>
			<% If rc("PrintPaidSum") = "N" Then %>
			<tr class="GeneralTbl">
				<td class="GeneralTblBold2"  width="124"><%=getfinalPaymentLngStr("LtxtPendBal")%></td>
				<td class="GeneralTbl" width="120" align="right"><nobr><%=rc("DocCur")%>&nbsp;<%=FormatNumber(TotalCancelado,myApp.SumDec)%></nobr></td>
			</tr>
			<% End If %>
		</table>
		</td>
	</tr>
	<tr>
		<td width="99%" colspan="2">
		<div align="left">
			<table border="0" cellpadding="0" width="275" id="table7">
				<tr class="GeneralTblBold2">
					<td><%=getfinalPaymentLngStr("DtxtObservations")%></td>
				</tr>
				<tr class="GeneralTbl">
					<td><% If Not IsNull(Rc("Comments")) Then %><%=Rc("Comments")%><% End If %>&nbsp;</td>
				</tr>
			</table>
		</div>
		</td>
	</tr>
	</table>
<% If rc("Confirm") = "N" and Session("NotifyAdd") and Request("Confirm") <> "Y" Then 
	Session("NotifyAdd") = False
	sql = "EXEC OLKCommon..DBOLKObjAlert" & Session("ID") & " " & Session("ConfPayRetVal") & ", " & Session("branch") & ", '" & userType & "', '" & getMyLng & "'"
	conn.execute(sql)
End If

Function getPtmntCurField
	If DocType = 17 and myApp.SVer = 5 Then
		getPtmntCurField = "DocCurr"
	Else
		getPtmntCurField = "DocCur"
	End If
End Function
%>
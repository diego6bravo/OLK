<%@ Language=VBScript %>
<% If session("OLKDB") = "" Then response.redirect "lock.asp" %>
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="lang/cxcDocDetail.asp" -->
<!--#include file="myHTMLEncode.asp"-->
<html <% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>>
<%
Dim varx
varx = "0"
%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" >
<title><%=getcxcDocDetailLngStr("LtxtDocDet")%></title>
<%
Dim DocName
           set rs = Server.CreateObject("ADODB.recordset")
           sql = "select top 1 DirectRate, SelDes, VatPrcnt " & _
           "from oadm cross join OLKCommon " & _
           "order by CurrPeriod desc"
           set rs = conn.execute(sql)
           DirectRate = rs("DirectRate")
           VatPrcnt = rs("VatPrcnt")
           If userType = "V" Then SelDes = 0 Else SelDes = rs("SelDes") %>
<!--#include file="loadAlterNames.asp" -->
<!--#include file="clearItem.asp" -->
<%      
          
           
           set rd = Server.CreateObject("ADODB.RecordSet")
           set rdocnum = Server.CreateObject("ADODB.recordset")
           If Request("isEntry") <> "N" Then isEntry = "Y" Else isEntry = "N"
           
           sql = "exec OLKCommon..DBOLKCXCDocInfo" & Session("ID") & " "  & Request.Form("doctype") & ", " & Request.Form("DocEntry") & ", " & Session("LanID") & ", '" & isEntry & "' "
           set rs = conn.execute(sql)
           
           If Not rs.Eof Then
           
           DocEntry = rs("DocEntry")
           If userType = "C" and LCase(saveHTMLDecode(Session("username"), False)) <> LCase(rs("CardCode")) and LCase(saveHTMLDecode(Session("username"), False)) <> LCase(rs("FatherCard")) Then Response.Redirect "noaccess.asp?ErrCode=0"

			If Request("DocType") <> "-2" Then
				If Request("DocType") <> "112" Then
					sql= "select T1, T2 from olkDocConf where ObjectCode = " & Request("DocType")
					set rd = conn.execute(sql)
					oTable = rd("T1")
					oTable1 = rd("T2")
				Else
					oTable = "ODRF"
					oTable1 = "DRF1"
				End If
			End If  
			
			If Request("DocType") = "-2" Then
				sqlAddStr = "(select SDKID collate database_default from r3_obscommon..tcif where companydb = N'" & Session("OlkDB") & "')"
			Else
				sqlAddStr = "'U_'"
			End If
			         
			set rcOpt = Server.CreateObject("ADODB.RecordSet")
			sql = "select T0.FieldID, AliasID, IsNull(AlterDescr, Descr) Descr, TypeID, EditType, SizeID, Dflt, NotNull, Pos, " & _
				  "Case When Exists(select 'A' from UFD1 where TableId = T0.TableId and FieldId = T0.FieldId) " & _
				  "Then 'Y' Else 'N' End As DropDown, NullField, Query, " & _
				  sqlAddStr & "++AliasID collate database_default As InsertID " & _
				  "from cufd T0 " & _
				  "left outer join OLKCUFD T1 on T1.TableID = T0.TableID and T1.FieldID = T0.FieldID " & _
				  "left outer join OLKCUFDAlterNames T2 on T2.TableID = T0.TableID and T2.FieldID = T0.FieldID and T2.LanID = " & Session("LanID") & " " & _
				  "left outer join OLKCUFDGroups T3 on T3.TableID = T1.TableID and T3.GroupID = IsNull(T1.GroupID, -1) " & _
				  "where T0.TableId = 'OINV' and AType in ('" & userType & "', 'T') and OP in ('T','P') and Active = 'Y' " & _
				  "order by T3.[Order], IsNull(T1.Pos, 'D'), IsNull(T1.[Order], 32727) "
			rcOpt.open sql, conn, 3, 1

			set rctn = Server.CreateObject("ADODB.RecordSet")
			sql = "select AliasID, NullField, IsNull(AlterDescr, Descr) Descr " & _
				  "from cufd T0 " & _
				  "left outer join OLKCUFD T1 on T1.TableID = T0.TableID and T1.FieldID = T0.FieldID " & _
				  "left outer join OLKCUFDAlterNames T2 on T2.TableID = T0.TableID and T2.FieldID = T0.FieldID and T2.LanID = " & Session("LanID") & " " & _
				  "where T0.TableId = 'OINV' and AType in ('" & userType & "', 'T') and OP in ('T','P') and Active = 'Y' and NullField = 'Y' " & _
				  "Order By Pos Desc"
			rctn.open sql, conn, 3, 1
			
			If rctn.recordcount > 0 Then chkOpt = True
			
			If rcOpt.RecordCount > 0 Then
				set rcOptVals = Server.CreateObject("ADODB.RecordSet")
				sql = "select " & sqlAddStr & "++AliasID As InsertID, T0.FieldID, " & _
					  "Case When Exists(select 'A' from UFD1 where TableId = T0.TableId and FieldId = T0.FieldId) " & _
				  	  "Then 'Y' Else 'N' End As DropDown, RTable " & _
				  	  "from cufd T0 " & _
				  	  "left outer join OLKCUFD T1 on T1.TableID = T0.TableID and T1.FieldID = T0.FieldID " & _
					  "where T0.TableId = 'OINV' and AType in ('" & userType & "','T') and OP in ('T','P') and Active = 'Y'"
				rcOptVals.open sql, conn, 3, 1
				sql = "select "
				do while not rcOptVals.eof
					If rcOptVals.bookmark <> 1 Then sql = sql & ", "
					If rcOptVals("DropDown") = "Y" and IsNull(rcOptVals("RTable")) Then
						sql = sql & "(select IsNull((select AlterDescr from OLKUFD1AlterNames where TableID = UFD1.TableID and FieldID = UFD1.FieldID and IndexID = UFD1.IndexID and LanID = " & Session("LanID") & "), Descr) from UFD1 where TableID = 'OINV' and FieldID = " & rcOptVals("FieldID") & " and FldValue = T0." & rcOptVals("InsertID") & " collate database_default) '" & rcOptVals("InsertID") & "'"
					ElseIf Not IsNull(rcOptVals("RTable")) Then
						sql = sql & "(select Name from [@" & rcOptVals("RTable") & "] where Code = T0." & rcOptVals("InsertID") & " collate database_default) '" & rcOptVals("InsertID") & "'"
					Else
						sql = sql & rcOptVals("InsertID")
					End If
				rcOptVals.movenext
				loop
				If Request("DocType") = "-2" Then
					sql = sql & " from r3_obscommon..tdoc T0 where lognum = " & DocEntry
				Else
					sql = sql & " from " & oTable & " T0 where T0.DocEntry = " & DocEntry
				End If
				set rcOptVals = conn.execute(sql)
			End If
			

			
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
				strContent = Replace(strContent, "{CmpName}", rs("CompnyName"))

           %>
<link rel="stylesheet" type="text/css" href="design/<%=SelDes%>/style/stylenuevo.css">
</head>
<body<% If Session("rtl") <> "" Then %> dir="rtl"<% End If %>>
<table border="0" cellpadding="0" width="100%" id="table1">
	<%=strContent%>
	<!--#include file="myTitle.asp"-->
	<tr class="CanastaTitle">
		<td>
		<p><strong><%=rs("ObjDesc")%> #<%=rs("DocNum")%></strong></td>
	</tr>
	<tr class="CanastaTitle2">
		<td>
		<table border="0" cellpadding="0" width="100%" cellspacing="1" id="table2">
			<tr>
				<td width="50%" valign="top">
				<table border="0" cellpadding="0" width="100%" id="table5">
					<tr>
						<td class="CanastaTblResaltada" width="95"><%=getcxcDocDetailLngStr("LtxtTo")%></td>
						<td class="CanastaTbl"><%=RS("cardname")%></td>
					</tr>
					<tr>
						<td class="CanastaTblResaltada" width="95" valign="top">
						<nobr><%=getcxcDocDetailLngStr("LtxtShipAdd")%></nobr></td>
						<td class="CanastaTbl">
						<%=RS("ShipAddress")%>&nbsp;</td>
					</tr>
					<tr>
						<td class="CanastaTblResaltada" width="95" valign="top"><%=getcxcDocDetailLngStr("LtxtPayAdd")%></td>
						<td class="CanastaTbl"><%=RS("PayAddress")%></td>
					</tr>
					<tr>
						<td class="CanastaTblResaltada" width="95"><%=getcxcDocDetailLngStr("DtxtPhone")%></td>
						<td class="CanastaTbl"><%=RS("Phone1")%></td>
					</tr>
					<tr>
						<td class="CanastaTblResaltada" width="95"><%=getcxcDocDetailLngStr("DtxtFax")%></td>
						<td class="CanastaTbl"><%=RS("fax")%></td>
					</tr>
					<tr>
						<td class="CanastaTblResaltada" width="95"><%=getcxcDocDetailLngStr("LtxtEMail")%></td>
						<td class="CanastaTbl"><%=RS("e_mail")%></td>
					</tr>
					<tr>
						<td class="CanastaTblResaltada" width="95">
						<%=getcxcDocDetailLngStr("DtxtContact")%></td>
						<td class="CanastaTbl">
						<%=RS("Name")%>&nbsp;&nbsp;</td>
					</tr>

					<% rcOpt.Filter = "Pos = 'I'"
					do while not rcOpt.eof 
                    AliasID = rcOpt("InsertID") %>
                    <tr>
                      <td class="CanastaTblResaltada" width="95">
                      <%=rcOpt("Descr")%>&nbsp;</td>
                      <td class="CanastaTbl">
                      <% If rcOpt("TypeID") = "M" and rcOpt("EditType") = "B" Then %><a class="LinkNoticiasMas" target="_blank" href="<%=rcOptVals(AliasID)%>"><% End If %>
                      <% If rcOpt("TypeID") = "B" Then
                      		If Not IsNull(rcOptVals(AliasID)) Then
			            	Select Case rcOpt("EditType")
								Case "R"
									Response.Write FormatNumber(CDbl(rcOptVals(AliasID)),myApp.RateDec)
								Case "S"
									Response.Write FormatNumber(CDbl(rcOptVals(AliasID)),myApp.SumDec)
								Case "P"
									Response.Write FormatNumber(CDbl(rcOptVals(AliasID)),myApp.PriceDec)
								Case "Q"
									Response.Write FormatNumber(CDbl(rcOptVals(AliasID)),myApp.QtyDec)
								Case "%"
									Response.Write FormatNumber(CDbl(rcOptVals(AliasID)),myApp.PercentDec)
								Case "M"
									Response.Write FormatNumber(CDbl(rcOptVals(AliasID)),myApp.MeasureDec)
			            	End Select
			            	End If
                      ElseIf rcOpt("TypeID") = "A" and rcOpt("EditType") = "I" Then %>
                      <img src="pic.aspx?filename=<% If Not IsNull(rcOptVals(AliasID)) Then %><%=rcOptVals(AliasID)%><% Else %>n_a.gif<% End If %>&MaxSize=180&dbName=<%=Session("OlkDb")%>" border="0">
                      <% Else %>
                      <% If Not IsNull(rcOptVals(AliasID)) Then %><%=rcOptVals(AliasID)%><% End If %>
                      <% End If %>
                      <% If rcOpt("TypeID") = "M" and rcOpt("EditType") = "B" Then %></a><% End If %>
                      </td>
                    </tr>
                    <% rcOpt.movenext
                    loop  %>
					</table>
				</td>
				<td valign="top" width="50%">
				<table border="0" cellpadding="0" width="100%" id="table7">
					<tr>
						<td class="CanastaTblResaltada" width="95"><%=getcxcDocDetailLngStr("DtxtDate")%></td>
						<td class="CanastaTbl">
						<%=rs("DocDate")%>&nbsp;</td>
					</tr>
					<tr>
						<td class="CanastaTblResaltada" width="95"><% Select Case rs("ObjType")
                      	Case 13
                      		Response.Write getcxcDocDetailLngStr("DtxtInvDueDate")
                      	Case 15
                      		Response.Write getcxcDocDetailLngStr("DtxtDeliveryDate")
                      	Case 17
                      		Response.Write getcxcDocDetailLngStr("DtxtDeliveryDate")
                      	Case 23
                      		Response.Write getcxcDocDetailLngStr("DtxtComDate")
                      	Case Else %><%=getcxcDocDetailLngStr("LtxtDueDate")%><% End Select %></td>
						<td class="CanastaTbl">
						<%=rs("DocDueDate")%>&nbsp;</td>
					</tr>
					<tr>
						<td class="CanastaTblResaltada" width="95"><% If 1 = 2 Then %>txtRef2<% Else %><%=Server.HTMLEncode(txtRef2)%><% End If %></td>
						<td class="CanastaTbl"><%=RS("NumAtCard")%>&nbsp;</td>
					</tr>
					<tr>
						<td class="CanastaTblResaltada" width="95"><% If 1 = 2 Then %>txtAgent<% Else %><%=Server.HTMLEncode(txtAgent)%><% End If %></td>
						<td class="CanastaTbl"><%=RS("SlpName")%>&nbsp;</td>
					</tr>
					<% rcOpt.Filter = "Pos = 'D'"
					do while not rcOpt.eof 
                    AliasID = rcOpt("InsertID") %>
                    <tr>
                      <td class="CanastaTblResaltada" width="95">
                      <%=rcOpt("Descr")%>&nbsp;</td>
                      <td class="CanastaTbl">
                      <% If rcOpt("TypeID") = "M" and rcOpt("EditType") = "B" Then %><a class="LinkNoticiasMas" target="_blank" href="<%=rcOptVals(AliasID)%>"><% End If %>
                      <% If rcOpt("TypeID") = "B" Then
                      		If Not IsNull(rcOptVals(AliasID)) Then
			            	Select Case rcOpt("EditType")
								Case "R"
									Response.Write FormatNumber(CDbl(rcOptVals(AliasID)),myApp.RateDec)
								Case "S"
									Response.Write FormatNumber(CDbl(rcOptVals(AliasID)),myApp.SumDec)
								Case "P"
									Response.Write FormatNumber(CDbl(rcOptVals(AliasID)),myApp.PriceDec)
								Case "Q"
									Response.Write FormatNumber(CDbl(rcOptVals(AliasID)),myApp.QtyDec)
								Case "%"
									Response.Write FormatNumber(CDbl(rcOptVals(AliasID)),myApp.PercentDec)
								Case "M"
									Response.Write FormatNumber(CDbl(rcOptVals(AliasID)),myApp.MeasureDec)
			            	End Select
			            	End If
                      ElseIf rcOpt("TypeID") = "A" and rcOpt("EditType") = "I" Then %>
                      <img src="pic.aspx?filename=<% If Not IsNull(rcOptVals(AliasID)) Then %><%=rcOptVals(AliasID)%><% Else %>n_a.gif<% End If %>&MaxSize=180&dbName=<%=Session("OlkDb")%>" border="0">
                      <% Else %>
                      <% If Not IsNull(rcOptVals(AliasID)) Then %><%=rcOptVals(AliasID)%><% End If %>
                      <% End If %>
                      <% If rcOpt("TypeID") = "M" and rcOpt("EditType") = "B" Then %></a><% End If %>
                      </td>
                    </tr>
                    <% rcOpt.movenext
                    loop %>
					</table>
				</td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table8">
			<tr class="CanastaTblResaltada">
				<td align="center">#</td>
				<% If myApp.GetShowRef Then %><td align="center"><%=getcxcDocDetailLngStr("DtxtCode")%></td><% End If %>
				<td align="center"><%=getcxcDocDetailLngStr("LtxtProd")%></td>
				<td align="center"><%=getcxcDocDetailLngStr("DtxtQty")%></td>
				<% If myApp.GetShowSalUn Then %><td align="center"><%=getcxcDocDetailLngStr("DtxtUnit")%></td><% End If %>
				<td align="center"><%=getcxcDocDetailLngStr("DtxtPrice")%></td>
				<td align="center"><%=getcxcDocDetailLngStr("DtxtTotal")%></td>
			</tr>
			  <% 
			  set cmd = Server.CreateObject("ADODB.Command")
			  cmd.ActiveConnection = connCommon
			  cmd.CommandType = &H0004
			  cmd.CommandText = "DBOLKGetDocLinesDetails" & Session("ID")
			  cmd.Parameters.Refresh()
			  cmd("@LanID") = Session("LanID")
			  cmd("@DocType") = Request("DocType")
			  cmd("@SlpCode") = Session("vendid")
			  If DocType = -2 Then
			  	cmd("@LogNum") = DocEntry
			  Else
			  	cmd("@DocEntry") = DocEntry
			  End If
			  set rd = Server.CreateObject("ADODB.RecordSet")
			  set rd = cmd.execute()
			  lNum = 0
			  do while not rd.eof 
			  lNum = lNum + 1
			  If Request("high") <> "" Then
				  If CStr(rd("LineNum")) = CStr(Request("high")) Then
				  	HighLight = True
				  Else
				  	HighLight = False
				  End If
			  Else
			  	HighLight = False
			  End If
			  
			  TreeType = rd("TreeType")
			  ShowPriceAndTotal = (TreeType <> "S" and TreeType <> "I") or ((TreeType = "S" or TreeType = "I") and rd("Currency") <> "" and CDbl(rd("LineTotal")) <> 0)
			  %>
			<tr class="CanastaTbl" <% If HighLight then %>style="color: #BB0000"<% End If %>>
				<td align="right"><%=lNum%>&nbsp;&nbsp;</td>
				<% If myApp.GetShowRef Then %><td><%=RD("ItemCode")%>&nbsp;</td><% End If %>
				<td><%=RD("ItemName")%>&nbsp;</td>
				<td>
				<p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">&nbsp;<%=FormatNumber(Cdbl(RD("Quantity")),myApp.QtyDec)%></td>
				<% If myApp.GetShowSalUn Then %><td align="center"><%=RD("SalUnitMsr")%>&nbsp;</td><% End If %>
				<td>
				<p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><% If ShowPriceAndTotal Then %><nobr><%=rd("Currency")%>&nbsp;<%=FormatNumber(RD("Price"),myApp.PriceDec)%></nobr><% Else %>&nbsp;<% End If %></td>
				<td>
				<p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><% If ShowPriceAndTotal Then %><nobr><%=Rs("DocCur")%>&nbsp;<%=FormatNumber(RD("LineTotal"),myApp.SumDec)%></nobr><% Else %>&nbsp;<% End If %></td>
			</tr>
		    <% Rd.MoveNext
			loop %>
			<tr>
				<td rowspan="8" valign="top">
				&nbsp;</td>
				<td colspan="<% If myApp.GetShowRef and myApp.GetShowSalUn Then %>4<% ElseIf myApp.GetShowRef or myApp.GetShowSalUn Then %>3<% Else %>2<% End If %>" rowspan="8" valign="top">
				<div align="left">
				<% If rs("PayLogNum") <> "" then 
				sql = "declare @LogNum int set @LogNum = " & rs("PayLogNum") & " " & _
					  "select DocCur, IsNull(CashSum,0) CashSum, (select IsNull(Sum(CheckSum),0) from r3_obscommon..pmt1 where lognum = @lognum) CheckSum, " & _
					  "(select IsNull(Sum(CreditSum),0) from r3_obscommon..pmt3 where lognum = @lognum) CreditSum, IsNull(TrsfrSum,0) TrsfrSum, " & _
					  "IsNull(CashSum,0)+ " & _
          	 		  "(select IsNull(Sum(CheckSum),0) from r3_obscommon..pmt1 where lognum = @lognum)+ " & _
	           		  "(select IsNull(Sum(CreditSum),0) from r3_obscommon..pmt3 where lognum = @lognum)+ " & _
       		 		  "IsNull(TrsfrSum,0) SumApplied " & _
       		 		  "from R3_ObsCommon..TPMT where LogNum = @LogNum "
				set rd = conn.execute(sql)
				SumApplied = rd("SumApplied")
				PayDocCur = rd("DocCur")
				%>
				<table border="0" cellpadding="0" width="80%" cellspacing="1" id="table6">
					<tr class="FirmTlt3">
						<td colspan="4"><% If 1 = 2 Then %>txtRct<% Else %><%=txtRct%><% End If %> #<font color="#BB0000"><%=rctNum%></font></td>
					</tr>
					<tr class="FirmTlt3">
						<td align="center"><%=getcxcDocDetailLngStr("LtxtCash")%></td>
						<td align="center"><%=getcxcDocDetailLngStr("LtxtCheck")%></td>
						<td align="center"><%=getcxcDocDetailLngStr("LtxtBankTrans")%></td>
						<td align="center"><%=getcxcDocDetailLngStr("LtxtCredCard")%></td>
					</tr>
					<tr class="FirmTbl">
						<td align="center"><nobr><%=PayDocCur%>&nbsp;<%=FormatNumber(Rd("CashSum"),myApp.SumDec)%></nobr></td>
						<td align="center"><nobr><%=PayDocCur%>&nbsp;<%=FormatNumber(Rd("CheckSum"),myApp.SumDec)%></nobr></td>
						<td align="center"><nobr><%=PayDocCur%>&nbsp;<%=FormatNumber(Rd("TrsfrSum"),myApp.SumDec)%></nobr></td>
						<td align="center"><nobr><%=PayDocCur%>&nbsp;<%=FormatNumber(Rd("CreditSum"),myApp.SumDec)%></nobr></td>
					</tr>
				</table>
				<% End If %>
				<br>

		<table border="0" cellpadding="0" width="100%" cellspacing="1" id="table9">
			<tr>
				<td width="261" valign="top">
				<table border="0" cellpadding="0" width="100%" id="table10">
					<tr class="CanastaTblResaltada">
						<td class=""><%=getcxcDocDetailLngStr("LtxtPymntCond")%></td>
					</tr>
					<tr class="CanastaTbl">
						<td><%=RS("PymntGroup")%>&nbsp;</td>
					</tr>
					<tr class="CanastaTblResaltada">
						<td><%=getcxcDocDetailLngStr("DtxtObservations")%></td>
					</tr>
					<tr class="CanastaTbl">
						<td>
						<%=RS("Comments")%>&nbsp;</td>
					</tr>
				</table>
				</td>
				<td valign="top">&nbsp;</td>
			</tr>
		</table></div>
				</td>
				<td class="CanastaTblResaltada"><%=getcxcDocDetailLngStr("LtxtSubTotal")%></td>
				<td class="CanastaTbl">
				<p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><nobr><%=rs("DocCur")%>&nbsp;<%=FormatNumber(Rs("SubTotal"),myApp.SumDec)%></nobr></td>
			</tr>
          <%
            If myApp.ExpItems Then
				cmd.CommandText = "DBOLKGetDocExpnsDetails" & Session("ID")
				cmd.Parameters.Refresh()
				cmd("@DocType") = Request("DocType")
				If DocType = -2 Then cmd("@LogNum") = DocEntry Else cmd("@DocEntry") = DocEntry
				rd.close
				rd.open cmd, , 3, 1
			  do while not rd.eof
			%>
			<tr>
				<td class="CanastaTblResaltada"><%=rd("ItemName")%>:&nbsp;</td>
				<td class="CanastaTbl">
				<p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
				<nobr><%=Rd("Currency")%>&nbsp;<%=FormatNumber(Rd("LineTotal"),myApp.SumDec)%></nobr></td>
			</tr>
			<% rd.movenext
			loop 
			End If %>
			<tr>
				<td class="CanastaTblResaltada"><%=getcxcDocDetailLngStr("DtxtDiscount")%></td>
				<td class="CanastaTbl">
				<p align="right" dir="ltr"><nobr>%&nbsp;<%=FormatNumber(CDbl(rs("DiscPrcnt")),myApp.SumDec)%></nobr></td>
			</tr>
			<tr>
				<td class="CanastaTblResaltada"><% If 1 = 2 Then %>txtTax<% Else %><%=txtTax%><% End If %><% If myApp.LawsSet = "IL" Then %><% If Session("myLng") = "he" Then %><span style="font-size: xx-small; "><% End If %><nobr>%&nbsp;<%=FormatNumber(VatPrcnt, myApp.PercentDec)%></nobr><% If Session("myLng") = "he" Then %></span><% End If %><% End If %></td>
				<td class="CanastaTbl">
				<p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><nobr><%=RS("DocCur")%>&nbsp;<%=FormatNumber(Rs("ITBM"),myApp.SumDec)%></nobr></td>
			</tr>
			<tr>
				<td class="CanastaTblResaltada"><%=getcxcDocDetailLngStr("DtxtTotal")%></td>
				<td class="CanastaTbl">
				<p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><nobr><%=RS("DocCur")%>&nbsp;<%=FormatNumber(Rs("DocTotal"),myApp.SumDec)%></nobr></td>
			</tr>
			<% If Not IsNull(rs("PayLogNum")) Then %>
			<tr>
				<td class="CanastaTblResaltada"><%=getcxcDocDetailLngStr("DtxtPaid")%></td>
				<td class="CanastaTbl">
				<nobr><%=PayDocCur%>&nbsp;<%=FormatNumber(SumApplied,myApp.SumDec)%></nobr></td>
			</tr>
			<tr>
				<td class="CanastaTblResaltada"><%=getcxcDocDetailLngStr("DtxtBalance")%></td>
				<td class="CanastaTbl" width="134">
				 <% OpenSum = CDbl(Rs("DocTotal"))-CDbl(SumApplied)
                  If OpenSum < 0 Then OpenSum = 0 %>
                  <nobr><%=RS("DocCur")%>&nbsp;<%=FormatNumber(OpenSum,myApp.SumDec)%></nobr></td>
			</tr>
			<% End If %>
			<tr>
				<td class="CanastaTblResaltada">&nbsp;</td>
				<td class="CanastaTbl" width="134">
				 &nbsp;</td>
			</tr>
			<%
			If Request("doctype") = "-2" Then
			set rx = Server.CreateObject("ADODB.RecordSet")
			sql = "select IsNull(T1.AlterRowName, T0.RowName) RowName, T0.RowQuery " & _
			"from OLKCMREP T0 " & _
			"left outer join OLKCMREPAlterNames T1 on T1.RowType = T0.RowType and T1.LineIndex = T0.LineIndex and T1.LanID = " & Session("LanID") & " " & _
			"where T0.RowActive = 'Y' and T0.Print" & userType & " = 'Y' " & _
			"order by T0.RowOrder asc"
			rx.open sql, conn, 3, 1
			if not rx.eof then
			sql = "declare @LogNum int set @LogNum = " & Request.Form("DocEntry") & " " & _
			"declare @CardCode nvarchar(15) set @CardCode = N'" & saveHTMLDecode(rs("CardCode"), False) & "' " & _
			"select "
			do while not rx.eof
				If rx.bookmark > 1 Then sql = sql & ", "
				sql = sql & "(" & rx("RowQuery") & ") As '" & rx("RowName") & "'"
			rx.movenext
			loop
			sql = QueryFunctions(sql)
			set rx = conn.execute(sql)
			For each fld in rx.Fields
			%>
			<tr>
				<td class="CanastaTblResaltada">&nbsp;</td>
				<td class="CanastaTblResaltada"><%=fld.Name%>&nbsp;</td>
				<td class="CanastaTbl" width="134" align="right">&nbsp;<%=fld%></td>
			</tr>
			<% Next %>
			<% End If
			End If %>
			</table>
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
</table>
<% Else %>
<%=getcxcDocDetailLngStr("DtxtNoData")%>
<% End If %></body><% set rs = nothing
set rd = nothing
set rx = Nothing
conn.close %></html>
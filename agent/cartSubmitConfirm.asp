<% addLngPathStr = "" %>
<!--#include file="lang/cartSubmitConfirm.asp" -->
<% 
ShowClientRef = myApp.ShowClientRef or userType = "V"

set rc = Server.CreateObject("ADODB.recordset")
set rs = Server.CreateObject("ADODB.recordset")
set rd = Server.CreateObject("ADODB.recordset")
set rs = Server.CreateObject("ADODB.recordset")

MainCur = myApp.MainCur
TreePricOn = myApp.TreePricOn
EnableDiscount = False
If userType = "V" Then
	EnableDiscount = myApp.EnableDiscount
	PrintPriceBefDiscount = myApp.PrintPriceBefDiscount
	PrintLineDiscount = myApp.PrintLineDiscount
End If

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetCartConfirmInfo" & Session("ID")
cmd.Parameters.Refresh()
cmd("@LogNum") = CLng(Session("ConfRetVal"))
cmd("@LanID") = CInt(Session("LanID"))
cmd("@MainCur") = myApp.MainCur
cmd("@LawsSet") = myApp.LawsSet
If Request("payment") = "Y" Then cmd("@PayLogNum") = Session("ConfPayRetVal")
set rs = cmd.execute()
DiscPrcnt = CDbl(rs("DiscPrcnt"))
DocDate = rs("DocDate")
CCartNote = rs("CCartNote")
If Request("payment") = "Y" Then PayDocCur = rs("PayDocCur")

DocCur = rs("DocCur")

set rc = Server.CreateObject("ADODB.RecordSet")
cmd.CommandText = "DBOLKGetUDFReadCols" & Session("ID")
cmd.Parameters.Refresh()
cmd("@LanID") = Session("LanID")
cmd("@TableID") = "OINV"
cmd("@UserType") = userType
cmd("@OP") = "O"
rc.open cmd, , 3, 1
			
set rx = Server.CreateObject("ADODB.RecordSet")
sql = "select T0.RowType, T0.LineIndex, IsNull(T1.AlterRowName, T0.RowName) RowName, RowQuery " & _
"from OLKCMREP T0 " & _
"left outer join OLKCMREPAlterNames T1 on T1.RowType = T0.RowType and T1.LineIndex = T0.LineIndex and T1.LanID = " & Session("LanID") & " " & _
"where T0.RowActive = 'Y' and T0.Print" & userType & " = 'Y' and RowQuery is not null " & _
"order by T0.RowOrder asc"
rx.open sql, conn, 3, 1

addColSpan = 0
If ShowClientRef Then addColSpan = addColSpan + 1
If myApp.GetShowSalUn Then addColSpan = addColSpan + 1
If EnableDiscount Then
	If PrintLineDiscount Then addColSpan = addColSpan + 1
	If PrintPriceBefDiscount Then addColSpan = addColSpan + 1
End If 
			%>
<table border="0" cellpadding="0" width="100%">
	<% If rs("PrintCmpPaper") = "Y" Then %>
	<tr class="FirmTlt">
		<td height="80">&nbsp;</td>
	</tr>
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
	strContent = Replace(strContent, "{CmpName}", rs("CmpName")) %>
	<%=strContent %>
	<!--#include file="myTitle.asp"-->
	<tr>
		<td id="tdMyTtl" class="TablasTituloSec">
		<%=getcartSubmitConfirmLngStr("LttlCartPurConf")%></td>
	</tr>
	<% If rs("Draft") = "Y" or rs("Confirmed") = "N" Then %>
	<tr>
		<td id="tdMyTtl" class="TablasTituloDraft">
		<% If rs("Draft") = "Y" Then %><%=getcartSubmitConfirmLngStr("LttlDraftNote")%><% ElseIf rs("Confirmed") = "N" Then %><%=getcartSubmitConfirmLngStr("LttlConfirmNote")%><% End If %></td>
	</tr>
	<% End If %>
	<% End If %>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr>
				<td valign="top" width="50%">
				<table border="0" cellpadding="0" width="100%" cellspacing="1">
					<tr>
                      <td class="DatosTltIn" colspan="2"><b>
                      <% 
                      Select Case rs("object")
                      			Case "17"
	                      			sqlTable = "ordr"
	                      			docName = txtOrdr '"Pedido"
	                      			AlertID = 2
	                      		Case "15"
	                      			sqlTable = "odln"
	                      			docName = txtOdln '"Entrega"
	                      			AlertID = 7
                      			Case "23"
	                      			sqlTable = "oqut"
	                      			docName = txtQuote '"Cotizaci�n"
	                      			AlertID = 0
	                      		Case "13"
	                      			sqlTable = "oinv"
	                      			docName = txtInv '"Factura"
	                      			AlertID = 1
                       	End Select
                       	
				   		%>&nbsp;<%=getcartSubmitConfirmLngStr("DtxtLogNum")%>&nbsp;<font color="#BB0000"><%=Session("ConfRetval")%></font></b></td>
                    </tr>
					<% If Session("cart") = "cart" Then %>
					<tr>
						<td class="DatosTltIn" width="95"><%=getcartSubmitConfirmLngStr("DtxtCode")%></td>
						<td class="FirmTbl"><%=RS("CardCode")%>&nbsp;</td>
					</tr>
					<tr>
						<td class="DatosTltIn" width="95"><%=getcartSubmitConfirmLngStr("DtxtFor")%></td>
						<td class="FirmTbl"><%=RS("cardname")%>&nbsp;</td>
					</tr>
					<tr>
						<td class="DatosTltIn" width="95" valign="top"><nobr><%=getcartSubmitConfirmLngStr("LtxtShipAdd")%></nobr></td>
						<td class="FirmTbl"><%=rs("ShipToCode")%><br><%=RS("ShipAddress")%>&nbsp;</td>
					</tr>
					<tr>
						<td class="DatosTltIn" width="95" valign="top"><nobr><%=getcartSubmitConfirmLngStr("LtxtPayAdd")%></nobr></td>
						<td class="FirmTbl"><%=rs("PayToCode")%><br><%=RS("PayAddress")%>&nbsp;</td>
					</tr>
					<tr>
						<td class="DatosTltIn" width="95"><%=getcartSubmitConfirmLngStr("DtxtPhone")%></td>
						<td class="FirmTbl"><%=RS("Phone1")%>&nbsp;</td>
					</tr>
					<tr>
						<td class="DatosTltIn" width="95"><%=getcartSubmitConfirmLngStr("DtxtFax")%></td>
						<td class="FirmTbl"><%=RS("fax")%>&nbsp;</td>
					</tr>
					<tr>
						<td class="DatosTltIn" width="95"><%=getcartSubmitConfirmLngStr("DtxtEMail")%></td>
						<td class="FirmTbl"><%=RS("e_mail")%>&nbsp;</td>
					</tr>
					<tr>
						<td class="DatosTltIn" width="95"><%=getcartSubmitConfirmLngStr("DtxtContact")%></td>
						<td class="FirmTbl"><%=RS("Name")%>&nbsp;</td>
					</tr>
					<% else %>
					<tr>
						<td class="DatosTltIn" width="95"><%=getcartSubmitConfirmLngStr("DtxtFor")%></td>
						<td class="FirmTbl"><%=RS("cardname2")%>&nbsp;</td>
					</tr>
					<tr>
						<td class="DatosTltIn" width="95" valign="top"><nobr><%=getcartSubmitConfirmLngStr("LtxtShipAdd")%></nobr></td>
						<td class="FirmTbl"><%=RS("ShipAddress")%>&nbsp;</td>
					</tr>
					<tr>
						<td class="DatosTltIn" width="95" valign="top"><nobr><%=getcartSubmitConfirmLngStr("LtxtPayAdd")%></nobr></td>
						<td class="FirmTbl"><%=RS("PayAddress")%>&nbsp;</td>
					</tr>
					<% end if %>
					<% rc.Filter = "Pos = 'I'"
					do while not rc.eof
                    AliasID = "U_" & rc("AliasID") %>
					<tr>
                      <td class="DatosTltIn" width="95"><%=rc("Descr")%></td>
                      <td class="FirmTbl">
                      <% If rc("TypeID") = "M" and rc("EditType") = "B" Then %><a class="LinkNoticiasMas" target="_blank" href="<%=rs(AliasID)%>"><% End If %>
                      <% If rc("TypeID") = "B" Then
                      		If Not IsNull(rs(AliasID)) Then
			            	Select Case rc("EditType")
								Case "R"
									Response.Write FormatNumber(CDbl(rs(AliasID)),myApp.RateDec)
								Case "S"
									Response.Write FormatNumber(CDbl(rs(AliasID)),myApp.SumDec)
								Case "P"
									Response.Write FormatNumber(CDbl(rs(AliasID)),myApp.PriceDec)
								Case "Q"
									Response.Write FormatNumber(CDbl(rs(AliasID)),myApp.QtyDec)
								Case "%"
									Response.Write FormatNumber(CDbl(rs(AliasID)),myApp.PercentDec)
								Case "M"
									Response.Write FormatNumber(CDbl(rs(AliasID)),myApp.MeasureDec)
			            	End Select
			            	End If
                      ElseIf rc("TypeID") = "A" and rc("EditType") = "I" Then %>
                      <img src='pic.aspx?filename=<% If Not IsNull(rs(AliasID)) Then %><%=rs(AliasID)%><% Else %>n_a.gif<% End If %>&amp;MaxSize=180&amp;dbName=<%=Session("olkdb")%>' border="0">
                      <% Else %>
                      <%=rs(AliasID)%>
                      <% End If %>
                      <% If rc("TypeID") = "M" and rc("EditType") = "B" Then %></a><% End If %>
                      </td>
                    </tr>
                    <% 
                    rc.movenext
                    loop 
                    %>
					<tr class="DatosTltIn">
						<td colspan="2" style="height: 40px"></td>
					</tr>
				</table>
				</td>
				<td valign="top" width="50%">
				<table border="0" cellpadding="0" width="100%" cellspacing="1">
					<tr>
						<td class="DatosTltIn" width="95"><%=getcartSubmitConfirmLngStr("DtxtDate")%></td>
						<td class="FirmTbl"><%=FormatDate(rs("DocDate"), True)%>&nbsp;</td>
					</tr>
					<tr>
						<td class="DatosTltIn" width="95"><% Select Case rs("Object")
                      	Case 13
                      		Response.Write getcartSubmitConfirmLngStr("DtxtInvDueDate")
                      	Case 15
                      		Response.Write getcartSubmitConfirmLngStr("DtxtDeliveryDate")
                      	Case 17
                      		Response.Write getcartSubmitConfirmLngStr("DtxtDeliveryDate")
                      	Case 23
                      		Response.Write getcartSubmitConfirmLngStr("DtxtComDate")
                      	End Select %></td>
						<td class="FirmTbl"><%=FormatDate(rs("DocDueDate"), True)%>&nbsp;</td>
					</tr>
					<tr>
						<td colspan="2">&nbsp;</td>
					</tr>
					<tr>
						<td class="DatosTltIn" width="95"><nobr><% If 1 = 2 Then %>Marca<% Else %><%=txtRef2%><% End If %></nobr></td>
						<td class="FirmTbl"><%=RS("NumAtCard")%>&nbsp;</td>
					</tr>
					<tr>
						<td class="DatosTltIn" width="95"><% If 1 = 2 Then %>Agente<% Else %><%=txtAgent%><% End If %></td>
						<td class="FirmTbl"><%=RS("SlpName")%>&nbsp;</td>
					</tr>
                    <% rc.Filter = "Pos = 'D'"
                    do while not rc.eof
                    AliasID = "U_" & rc("AliasID") %>
					<tr>
                      <td class="DatosTltIn" width="95"><%=rc("Descr")%>&nbsp;</td>
                      <td class="FirmTbl">
                      <% If rc("TypeID") = "M" and rc("EditType") = "B" Then %><a class="LinkNoticiasMas" target="_blank" href="<%=rs(AliasID)%>"><% End If %>
                      <% If rc("TypeID") = "B" Then
                      		If Not IsNull(rs(AliasID)) Then
			            	Select Case rc("EditType")
								Case "R"
									Response.Write FormatNumber(CDbl(rs(AliasID)),myApp.RateDec)
								Case "S"
									Response.Write FormatNumber(CDbl(rs(AliasID)),myApp.SumDec)
								Case "P"
									Response.Write FormatNumber(CDbl(rs(AliasID)),myApp.PriceDec)
								Case "Q"
									Response.Write FormatNumber(CDbl(rs(AliasID)),myApp.QtyDec)
								Case "%"
									Response.Write FormatNumber(CDbl(rs(AliasID)),myApp.PercentDec)
								Case "M"
									Response.Write FormatNumber(CDbl(rs(AliasID)),myApp.MeasureDec)
			            	End Select
			            	End If
                      ElseIf rc("TypeID") = "A" and rc("EditType") = "I" Then %>
                      <img src="<%=cartPDFAddStr%>pic.aspx?filename=<% If Not IsNull(rs(AliasID)) Then %><%=rs(AliasID)%><% Else %>n_a.gif<% End If %>&MaxSize=180&dbName=<%=Session("olkdb")%>" border="0">
                      <% ElseIf rc("TypeID") = "D" Then %>
                      <%=FormatDate(rs(AliasID), True)%>
                      <% Else %>
                      <%=rs(AliasID)%>
                      <% End If %>
                      <% If rc("TypeID") = "M" and rc("EditType") = "B" Then %></a><% End If %>
                      </td>
                    </tr>
                    <% 
                    rc.movenext
                    loop 
                    %>
					<tr>
						<td class="DatosTltIn" width="95">&nbsp;</td>
						<td class="FirmTbl">&nbsp;</td>
					</tr>
					</table>
				</td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" cellspacing="1">
			<tr class="FirmTlt3">
				<% If ShowClientRef Then %><td align="center"><%=getcartSubmitConfirmLngStr("DtxtCode")%></td><% End If %>
				<td align="center"><%=getcartSubmitConfirmLngStr("DtxtDescription")%></td>
				<td align="center"><%=getcartSubmitConfirmLngStr("DtxtQty")%></td>
				<% If myApp.GetShowSalUn Then %><td align="center"><%=getcartSubmitConfirmLngStr("DtxtUnit")%></td><% End If %>
				<% If EnableDiscount Then %>
				<% If PrintPriceBefDiscount Then %>
				<td align="center" width="91"><%=getcartSubmitConfirmLngStr("LtxtUnitPrice")%></td>
				<% End If %>
				<% If PrintLineDiscount Then %>
				<td align="center" width="91"><%=getcartSubmitConfirmLngStr("DtxtDiscount")%></td>
				<% End If %>
				<% End If %>
				<td align="center" width="91"><% If Not EnableDiscount or EnableDiscount and Not (PrintPriceBefDiscount or PrintLineDiscount) Then %><%=getcartSubmitConfirmLngStr("DtxtPrice")%><% Else %><%=getcartSubmitConfirmLngStr("LtxtPriceAfterDisc")%><% End If %></td>
				<td align="center" width="134"><%=getcartSubmitConfirmLngStr("DtxtTotal")%></td>
			</tr>
			  <% 
			  If RS("Verfy") = "True" Then
				set cmd = Server.CreateObject("ADODB.Command")
				cmd.ActiveConnection = connCommon
				cmd.CommandType = &H0004
				cmd.CommandText = "DBOLKGetLinesInfo" & Session("ID")
				cmd.Parameters.Refresh()
				cmd("@LanID") = Session("LanID")
				cmd("@LogNum") = Session("ConfRetVal")
				cmd("@MainCur") = myApp.MainCur
				cmd("@SumDec") = myApp.SumDec
				cmd("@DirectRate") = GetYN(myApp.DirectRate)
				cmd("@LawsSet")= myApp.LawsSet
				cmd("@OptProm") = GetYN(optProm)
				cmd("@UserType") = userType
				cmd("@CardCode") = Session("UserName")
				cmd("@Enable3dx") = GetYN(myApp.Enable3dx)
				cmd("@SDKLineMemo") = GetYN(myApp.SDKLineMemo)
				cmd("@Object") = rs("Object")
				cmd("@PriceList") = Session("PriceList")
				cmd("@SlpCode") = Session("vendid")
				set rd = cmd.execute()

			  do While NOT Rd.EOF
			  
			  TreePricOn = rd("TreePricOn") = "Y"
			  TreeType = rd("TreeType")
			  ShowPriceAndTotal = Not (TreeType = "S" and Not TreePricOn or TreeType = "C" and (TreePricOn or Not IsNull(rd("ShowCompPrice")))) or TreeType = "S" and rd("ShowFatherPrice") = "Y" or TreeType = "C" and rd("ShowCompPrice") = "Y"
			  
			  %>
			<tr class="FirmTbl" style="<% Select Case TreeType
				Case "S" %>font-weight: bold;<%
				Case "C" %>font-style: italic;<%
			End Select %>">
				<% If ShowClientRef Then %><td><%=RD("ItemCode")%>&nbsp;</td><% End If %>
				<td><% If Not IsNull(RD("ItemName")) Then %><%=RD("ItemName")%><% End If %>&nbsp;</td>
				<td><p align="right">&nbsp;<%=FormatNumber(RD("Quantity"),myApp.QtyDec)%></td>
				<% If myApp.GetShowSalUn Then %><td align="center"><% Select Case rd("SaleType")
		  			Case 1
		  				Response.write getcartSubmitConfirmLngStr("DtxtUnit")
		  			Case 2
		  				Response.write rd("SalUnitMsr") 
		  				If myApp.GetShowQtyInUn Then Response.write "(" & rd("NumInSale") & ")"
		  			Case 3
					  	Response.write rd("SalPackMsr") 
					  	If myApp.GetShowQtyInUn Then Response.write "(" & rd("SalPackUn") & ")"
				  		If myApp.UnEmbPriceSet Then
					  		Response.write " x " & rd("SalUnitMsr") 
					  		If myApp.GetShowQtyInUn Then Response.write "(" & rd("NumInSale") & ")"
				  		End If
		  			End Select
			    %></td><% End If %>
				<% If EnableDiscount Then %>
				<% If PrintPriceBefDiscount Then %>
				<td width="91">
				<p align="right" dir="ltr"><% If ShowPriceAndTotal Then %><nobr><%=rd("Currency")%>&nbsp;<%=FormatNumber(RD("UnitPrice"),myApp.PriceDec)%></nobr><% Else %>&nbsp;<% End If %></td>
				<% End If %>
				<% If PrintLineDiscount Then %>
				<td width="91">
				<p align="right" dir="ltr"><% If ShowPriceAndTotal Then %><nobr>%&nbsp;<%=FormatNumber(RD("DiscPrcnt"),myApp.PercentDec)%></nobr><% End If %></td>
				<% End If %>
				<% End If %>
				<td width="91">
				<p align="right" dir="ltr"><% If ShowPriceAndTotal Then %><nobr><%=rd("Currency")%>&nbsp;<%=FormatNumber(RD("Price"),myApp.PriceDec)%></nobr><% Else %>&nbsp;<% End If %></td>
				<td width="134" align="right" dir="ltr"><% If ShowPriceAndTotal Then %><nobr><%=DocCur%>&nbsp;<%=FormatNumber(RD("LineTotal"),myApp.SumDec)%></nobr><% Else %>&nbsp;<% End If %></td>
			</tr>
		  <% 
		   Rd.MoveNext
		   loop 
		   end if %>
			<tr>
				<% ExpRowSpan = 7 + rx.recordcount
				If Request("payment") = "Y" Then
						ExpRowSpan = ExpRowSpan + rs("ETotal")
				End If %>
				<td colspan="<%=2+addColSpan%>" rowspan="<%=ExpRowSpan%>" valign="top">
				<div>
				<% If Request("payment") = "Y" then 
				%>
				<br>
				<table border="0" cellpadding="0" width="80%" cellspacing="1">
					<tr class="FirmTlt3">
						<td colspan="4"><% If 1 = 2 Then %><%=getcartSubmitConfirmLngStr("DtxtReceipt")%><% Else %><%=txtRct%><% End If %></td>
					</tr>
					<tr class="FirmTlt3">
						<td align="center"><%=getcartSubmitConfirmLngStr("LtxtCash")%></td>
						<td align="center"><%=getcartSubmitConfirmLngStr("LtxtCheques")%></td>
						<td align="center"><%=getcartSubmitConfirmLngStr("LtxtBankTran")%></td>
						<td align="center"><%=getcartSubmitConfirmLngStr("LtxtCredCards")%></td>
					</tr>
					<tr class="FirmTbl">
						<td align="center" dir="ltr" style="height: 20px"><nobr><%=PayDocCur%>&nbsp;<%=FormatNumber(rs("CashSum"),myApp.SumDec)%></nobr></td>
						<td align="center" dir="ltr" style="height: 20px"><nobr><%=PayDocCur%>&nbsp;<%=FormatNumber(rs("CheckSum"),myApp.SumDec)%></nobr></td>
						<td align="center" dir="ltr" style="height: 20px"><nobr><%=PayDocCur%>&nbsp;<%=FormatNumber(rs("TrsfrSum"),myApp.SumDec)%></nobr></td>
						<td align="center" dir="ltr" style="height: 20px"><nobr><%=PayDocCur%>&nbsp;<%=FormatNumber(rs("CreditSum"),myApp.SumDec)%></nobr></td>
					</tr>
				</table>
				<% end if %>
				<br>
				<table border="0" cellpadding="0" width="95%" cellspacing="1">
					<tr>
						<td class="DatosTltIn"><%=getcartSubmitConfirmLngStr("LtxtPymntGroup")%></td>
						<td class="FirmTbl"><%=RS("PymntGroup")%>&nbsp;</td>
					</tr>
					<tr>
						<td class="DatosTltIn" valign="top"><%=getcartSubmitConfirmLngStr("DtxtNote")%></td>
						<td class="FirmTbl"><%=CCartNote%>&nbsp;</td>
					</tr>
					<tr>
						<td class="DatosTltIn" valign="top"><%=getcartSubmitConfirmLngStr("LtxtDelObs")%></td>
						<td class="FirmTbl"><%=RS("Comments")%>&nbsp;</td>
					</tr>
				</table>
				</div>
				</td>
				<td class="DatosTltIn" width="91"><%=getcartSubmitConfirmLngStr("LtxtSubtotal")%></td>
				<td class="FirmTbl" width="134" align="right" dir="ltr"><nobr><%=DocCur%>&nbsp;<%=FormatNumber(CDbl(rs("SubTotal")),myApp.SumDec)%></nobr></td>
			</tr>
			<%  
			If myApp.ExpItems Then	
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKGetCartExpenses" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@LanID") = Session("LanID")
			cmd("@LogNum") = Session("ConfRetVal")
            set rd = cmd.execute()
            do while not rd.eof %>
			<tr>
				<td class="FirmTblY" width="91" style="font-weight: bold;"><% If Not IsNull(RD("ItemName")) Then %><%=rd("ItemName")%><% End If %>&nbsp;</td>
				<td class="FirmTblY" width="134" align="right"><nobr><%=DocCur%>&nbsp;<%=FormatNumber(Rd("Price"),myApp.SumDec)%></nobr></td>
			</tr>
            <% rd.movenext
            loop 
            End If
            If userType = "V" or userType = "C" and DiscPrcnt <> 0 Then %>
			<tr>
				<td class="DatosTltIn" width="91">
				<table cellpadding="0" cellspacing="0" border="0" width="100%">
					<tr class="DatosTltIn">
						<td><nobr><%=getcartSubmitConfirmLngStr("DtxtDiscount")%></nobr></td>
						<td style="padding-left: 2px; font-weight: normal;"><nobr><%=FormatNumber(DiscPrcnt,myApp.PercentDec)%>%</nobr></td>
					</tr>
				</table>
				</td>
				<td class="FirmTbl" width="134" align="right" dir="ltr"><nobr><%=DocCur%>&nbsp;<%=FormatNumber((CDbl(rs("SubTotal"))*(DiscPrcnt/100)),myApp.SumDec)%></nobr></td>
			</tr>
			<% End If %>
			<tr>
				<td class="DatosTltIn" width="91"><% If 1 = 2 Then %>Impuesto<% Else %><%=txtTax%><% End If %><% If myApp.LawsSet = "IL" Then %><% If Session("myLng") = "he" Then %><span style="font-size: xx-small; "><% End If %>&nbsp;%<%=FormatNumber(myApp.VatPrcnt, myApp.PercentDec)%><% If Session("myLng") = "he" Then %></span><% End If %><% End If %></td>
				<td class="FirmTbl" width="134" align="right" dir="ltr"><nobr><%=DocCur%>&nbsp;<%=FormatNumber(CDbl(rs("ITBM")),myApp.SumDec)%></nobr></td>
			</tr>
			<tr>
				<td class="DatosTltIn" width="91"><%=getcartSubmitConfirmLngStr("DtxtTotal")%></td><% varTotal = CDbl(rs("DocTotal")) %>
				<td class="FirmTbl" width="134" align="right" dir="ltr"><nobr><%=DocCur%>&nbsp;<%=FormatNumber(varTotal,myApp.SumDec)%></nobr></td>
			</tr>
			<% If Request("payment") = "Y" Then %>
			<tr>
				<td class="DatosTltIn" width="91"><%=getcartSubmitConfirmLngStr("LtxtPaid")%></td><% SumApplied = CDbl(rs("CashSum")) + CDbl(rs("CheckSum")) + CDbl(rs("CreditSum")) + CDbl(rs("TrsfrSum")) %>
				<td class="FirmTbl" width="134" align="right" dir="ltr"><nobr><%=PayDocCur%>&nbsp;<%=FormatNumber(SumApplied,myApp.SumDec)%></nobr></td>
			</tr>
			<% OpenSum = varTotal-SumApplied
                If OpenSum < 0 Then OpenSum = 0 %>
			<tr>
				<td class="DatosTltIn" width="91"><%=getcartSubmitConfirmLngStr("DtxtBalance")%></td>
				<td class="FirmTbl" width="134" align="right" dir="ltr"><nobr><%=DocCur%>&nbsp;<%=FormatNumber(OpenSum,myApp.SumDec)%></nobr></td>
			</tr>
			<% End If %>
			<tr>
				<td width="91">&nbsp;</td>
				<td width="134" align="right">&nbsp;</td>
			</tr>
			<% 
			if not rx.eof then
			sql = "declare @LogNum int set @LogNum = " & Session("ConfRetVal") & " " & _
			"declare @CardCode nvarchar(15) set @CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "' " & _
			"declare @LanID int set @LanID = " & Session("LanID") & " " & _
			"select "
			do while not rx.eof
				If rx.bookmark > 1 Then sql = sql & ", "
				sql = sql & "(" & rx("RowQuery") & ") As '" & rx("RowName") & "'"
			rx.movenext
			loop
			sql = QueryFunctions(sql)
			set rx = conn.execute(sql)
			For each fld in rx.Fields %>
			<tr>
				<td class="FirmTlt3" width="91"><%=fld.Name%></td>
				<td class="FirmTbl" width="134" align="right">&nbsp;<%=fld%></td>
			</tr>
			<% Next %>
			<% End If %>
			</table>
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
</table>
<% 
set rs = nothing
set rctn = nothing
set rd = nothing
 %>
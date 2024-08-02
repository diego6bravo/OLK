<% addLngPathStr = "cart/" %>
<!--#include file="lang/addCartMulti.asp" -->
<!--#include file="../itemFunctions.asp"-->
<% 
foundErr = False
globalErr = False
If Request("chkItem") <> "" Then
	set rs = Server.CreateObject("ADODB.Recordset")
	sql = 	"select Object from R3_ObsCommon..TLOG where LogNum = " & Session("RetVal") 
	set rs = conn.execute(sql)
	
	ChkInv = rs("Object") <> 23 
	
	arrItems = Split(Request("chkItem"), ", ")
	For i = 0 to UBound(arrItems)
		foundErr = False 
		If Request("resend") <> "Y" Then
			ItemCode = Request("chkItemCode" & arrItems(i))
			Qty = getNumeric(Request("chkItemQty" & arrItems(i)))
		Else
			resendID = arrItems(i)
			ItemCode = Request("ItemCode" & resendID)
			Qty = Request("Qty" & resendID)
		End If
		AddErr = getAddItmMultiError(ItemCode)
		If CStr(AddErr) <> "" Then
			foundErr = True
			globalErr = True
			arrItems(i) = ItemCode & "{S}Flow{S}" & AddErr & "{S}" & Qty
		End If
		
		If ChkInv Then
			sql = 	"declare @ItemCode nvarchar(20) set @ItemCode = N'" & saveHTMLDecode(ItemCode, False) & "' " & _
					"declare @whscode nvarchar(8) set @whscode = OLKCommon.dbo.DBOLKGetWhsCode" & Session("ID") & "(" & Session("branch") & ", " & Session("vendid") & ", @ItemCode) " & _
					"declare @SaleType2 int set @SaleType2 = " & myApp.GetSaleUnit & " " & _
					"declare @FirstQuantity numeric(19,6) set @FirstQuantity = " & Qty & " " & _
					"declare @UnEmbPriceSet char(1) set @UnEmbPriceSet = (select UnEmbPriceSet from olkcommon) " & _
					"declare @Quantity numeric(19,6) " & _
					"If @SaleType2 = 1 Begin  " & _
					"set @Quantity = @FirstQuantity/(select NumInSale from oitm where itemcode = @ItemCode) End  " & _
					"Else If @SaleType2 = 2 or @SaleType2 = 3 and @UnEmbPriceSet = 'Y' Begin  " & _
					"set @Quantity = @FirstQuantity End  " & _
					"Else If @SaleType2 = 3 and @UnEmbPriceSet = 'N' Begin  " & _
					"set @Quantity = @FirstQuantity*(select SalPackUn from oitm where itemcode = @ItemCode) End  " & _
					"select OLKCommon.dbo.DBOLKItemInv" & Session("ID") & "(@ItemCode, @WhsCode, @Quantity, '" & Session("olkdb") & "', -1, -1) Verfy"
			set rs = conn.execute(sql)
			If rs("Verfy") <> "Y" Then
				foundErr = True
				globalErr = True
				arrItems(i) = ItemCode & "{S}Inv{S}{S}" & Qty
			End If
		End If
		
		If not foundErr Then
			sql = 	"declare @ItemCode nvarchar(20) set @ItemCode = N'" & saveHTMLDecode(ItemCode, False) & "' " & _
					"declare @whscode nvarchar(8) set @whscode = OLKCommon.dbo.DBOLKGetWhsCode" & Session("ID") & "(" & Session("branch") & ", " & Session("vendid") & ", @ItemCode) " & _
					"select (select case when exists(select itemcode from r3_obscommon..doc1 where lognum = " & Session("RetVal") & _
					" and ItemCode = @ItemCode and whscode = @whscode) Then 'True' ELSE 'False' END) as Verfy, ISNULL((select Max(LineNum) from r3_obscommon..doc1 " & _
					"where LogNum = " & Session("RetVal") & " and ItemCode = @ItemCode and whscode = @whscode),0) As LineNum, " & _
					"ISNULL((select max(linenum)+1 from r3_obscommon..doc1 where lognum = " & Session("RetVal") & "),0) As MaxLineNum"
			set rs = conn.execute(sql)
			'no fue agregado lo agrega
			If rs("Verfy") = "False" or rs("Verfy") = "True" and myApp.BasketMItems Then
			      
			      
				set cmd = Server.CreateObject("ADODB.Command")
				cmd.ActiveConnection = connCommon
				cmd.CommandType = &H0004
				cmd.CommandText = "DBOLKCartAddSFM" & Session("ID")
				cmd.Parameters.Refresh()
				cmd("@lognum") = Session("RetVal")
				cmd("@quantity") = CDbl(Qty)
				cmd("@item") = ItemCode
				cmd("@PriceList") = Session("PList")
				cmd("@UserType") = "V"
				cmd("@SaleType") = myApp.GetSaleUnit
				cmd("@branchIndex") = Session("branch")
				cmd("@SlpCode") = Session("vendid")

				Select Case myApp.LawsSet 
					Case "MX", "CL", "CR", "GT", "US", "CA", "BR"
						If Request("TaxCode") <> "" Then
							TaxCode = Request("TaxCode")
						Else
							TaxCode = getItemTaxCode(ItemCode)
						End If
						
						If TaxCode = "Disabled" Then 
							TaxCode = "NULL"
						End If
						
						If TaxCode <> "NULL" Then cmd("@TaxCode") = TaxCode
				End Select				
				
				cmd.execute()
			'si ya fue agregado se le suma la cantidad que quiere comprar a la que ya fue agregada.
			ElseIf rs("Verfy") = "True" Then
				sql = ""
				If myApp.GetSaleUnit = 3 Then
					sql = "*(select SalPackUn from oitm where itemcode = N'" & saveHTMLDecode(ItemCode, False) & "') "
				End If
				sql = "update r3_obscommon..doc1 set quantity = quantity + " & Qty & " " & sql & " where LogNum = " & Session("RetVal") & " and linenum = " & rs("LineNum")
				conn.execute(sql)
			End If
			arrItems(i) = ItemCode & "{S}Add{S}{S}" & Qty
		End If
			
	Next
End If

If Not globalErr Then
	If myApp.AfterCartAddPocket = "Y" Then 
		response.redirect "operaciones.asp?cmd=slistsearch" 
	Else 
		response.redirect "operaciones.asp?cmd=cart"
	End If
Else
	%>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
  	<form name="frmConfirm" action="operaciones.asp" method="post">
    <tr>
      <td>
      <img src="images/spacer.gif" width="100%" height="1" border="0" alt></td>
    </tr>
    <tr>
      <td bgcolor="#9BC4FF">
      <table border="0" cellpadding="0" bordercolor="#111111" width="100%">
        <tr>
          <td width="100%">
          <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><%=getaddCartMultiLngStr("LttlAddItmRes")%></font></b></td>
        </tr>
        <tr>
          <td width="100%">
          <table border="0" cellpadding="0"  bordercolor="#111111" width="100%">
            <tr>
              <td bgcolor="#7DB1FF" align="center"><b><font size="1" face="Verdana"><%=getaddCartMultiLngStr("DtxtItem")%> - <%=getaddCartMultiLngStr("DtxtDescription")%></font></b></td>
              <td bgcolor="#7DB1FF" align="center" width="100"><b><font size="1" face="Verdana"><%=getaddCartMultiLngStr("DtxtState2")%></font></b></td>
            </tr>
            <%
            For i = 0 to UBound(arrItems)
            	arrData = Split(arrItems(i), "{S}")
            	If UBound(arrData) = 0 Then
            		lineStatus = "Add"
            	Else
            		lineStatus = arrData(1)
            	End If
				ItemCode = arrData(0)
				sql = "select OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'ItemCode', ItemCode, ItemName) " & _
						"from OITM " & _
						"where ItemCode = N'" & saveHTMLDecode(ItemCode, False) & "'"
				set rs = conn.execute(sql)
				ItemName = rs(0)
				Qty = arrData(3) %>
            <tr>
              <td bgcolor="#8CBAFF"><b><font size="1" face="Verdana"><%=ItemCode%> - <%=ItemName%></font></b></td>
              <td bgcolor="<% Select Case lineStatus
              Case "Add" %>#8CBAFF<% Case "Inv", "Flow" %>#FFD2A6<% End Select %>"><b><font size="1" face="Verdana"><% Select Case lineStatus
              															Case "Add" %><%=getaddCartMultiLngStr("LtxtAdded")%><%
              															Case "Inv" %><%=getaddCartMultiLngStr("LtxtInvErr")%><%
              															Case "Flow" %><%=getaddCartMultiLngStr("DtxtFlow")%><%
              														End Select %></font></b></td>
            </tr>
            <% If lineStatus = "Flow" Then
				rs.close
				sql = 	"select FlowID, Name, Type, " & _
						"ExecAt, Case When LineQuery is not null Then 'Y' Else 'N' End LineQry, NoteBuilder, NoteText, Draft, Authorize " & _
						"from OLKUAF " & _
						"where FlowID in (" & arrData(2) & ") " & _
						"order by Type asc, [Order] asc"
				rs.open sql, conn, 3, 1 %>
			<tr>
				<td colspan="2">
					<table border="0" cellpadding="0" width="100%">
						<% do while not rs.eof
							ExecAt = rs("ExecAt") %>
						<tr bgcolor="#7DB1FF">
							<td colspan="2" bgcolor="#7DB1FF">
							<font face="Verdana" size="1"><b><%=rs("Name")%></b>&nbsp;</font></td>
						</tr>
						<tr class="GeneralTbl">
							<td width="83%" valign="top"><font face="Verdana" size="1"><%=BuildMultiNote(rs("ExecAt"), ItemCode, Qty)%>&nbsp;</font></td>
							<td width="17%" valign="top">
							<p align="center">
							<% Select Case rs("Type")
								Case 1 %>
							<img border="0" src="images/questionicon.gif" width="68" height="65" alt="|L:altConf|">
							<%	Case 0 %>
							<img border="0" src="images/erroricon.gif" width="68" height="65" alt="|L:altError|">
							<% End Select %>
							</td>
						</tr>
						<% If rs("LineQry") = "Y" Then %>
						<tr class="GeneralTbl">
							<td colspan="2">
							<p align="center">
							<iframe name="content" width="100%" src='flowAlertDetails.asp?FlowID=<%=rs("FlowID")%>&amp;ExecAt=D2&amp;Item=<%=Server.URLEncode(ItemCode)%>&Price=<%=ItemPrice%>&Quantity=<%=Qty%>' border="0" frameborder="0" height="103">
							Your browser does not support inline frames or is currently configured not to display inline frames.
							</iframe></td>
						</tr>
						<% End If
						myType = rs("Type")
						rs.movenext
						loop
						If myType = 1 Then %>
						<tr class="GeneralTbl">
							<td colspan="2">
							<input type="checkbox" name="chkItem" id="chkItem<%=i%>" value="<%=i%>">
							<label for="chkItem<%=i%>"><font face="Verdana" size="1"><%=getaddCartMultiLngStr("DtxtConfirm")%></font></label>
							<input type="hidden" name="ItemCode<%=i%>" value="<%=Server.HTMLEncode(ItemCode)%>">
							<input type="hidden" name="DocConf<%=i%>" value="<%=arrData(2)%>">
							<input type="hidden" name="Qty<%=i%>" value="<%=Qty%>"></td>
						</tr><%
						End If %>
					</table>
				</td>
			</tr>
			<% End If %>
			<tr>
				<td colspan="2"><hr size="1"></td>
			</tr>
            <% Next %>
          </table>
          </td>
        </tr>
       </table>
       </td>
      </tr>
	  <tr>
		<td align="center"><input type="submit" name="btnAccept" value="<%=getaddCartMultiLngStr("DtxtAccept")%>"></td>
	  </tr>
		<input type="hidden" name="cmd" value="cartAddMulti">
		<input type="hidden" name="resend" value="Y">
      </form>
     </table>
  </center>
</div>

<% 
End If

Function getAddItmMultiError(Item)
	RetVal = ""

	set rFlow = Server.CreateObject("ADODB.RecordSet")
	set rChk = Server.CreateObject("ADODB.RecordSet")
	
	sqlFlow = 	"declare @ObjectCode int set @ObjectCode = (select Object from R3_ObsCommon..TLOG where LogNum = " & Session("RetVal") & ") " & _
				"select T0.FlowID, T0.Name, Type " & _
				"from OLKUAF T0  " & _
				"inner join OLKUAF1 T1 on T1.FlowID = T0.FlowID and T1.SlpCode in (" & Session("vendid") & ",-999) " & _
				"inner join OLKUAF2 T2 on T2.FlowID = T0.FlowID " & _
				"where T2.ObjectCode = @ObjectCode and T0.Active = 'Y' and T0.ExecAt = 'D2' "
	
	If Request("resend") = "Y" Then
		If Request("DocConf" & resendID) <> "" Then sqlFlow = sqlFlow & " and T0.FlowID not in (" & Request("DocConf" & resendID) & ") "
	End If
	
	sqlFlow = sqlFlow & " order by Type, [Order] asc"
	
	set rFlow = conn.execute(sqlFlow)

	do while not rFlow.eof
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKCheckDF" & Session("ID") & "_" & Replace(rFlow("FlowID"), "-", "_")
		LoadMultiParams
		set rChk = cmd.execute()
		If not rChk.eof then
			If Not IsNull(rChk(0)) Then
				If lcase(rChk(0)) = lcase("True") Then
					If RetVal <> "" Then RetVal = RetVal & ", "
					RetVal = RetVal & rFlow("FlowID")
					If rFlow("Type") = 0 Then Exit do
				End If
			End If
		End If
	rFlow.movenext
	loop
	getAddItmMultiError = RetVal
End Function

Function BuildMultiNote(ByVal ExecAt, ByVal ItemCode, ByVal Qty)
	myNote = rs("NoteText")
	If rs("NoteBuilder") = "Y" Then
	set cmd = Server.CreateObject("ADODB.Command")
		set rNote = Server.CreateObject("ADODB.RecordSet")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKCheckDF" & Session("ID") & "_" & Replace(rs("FlowID"), "-", "_") & "_msg"
		LoadMultiParams
		set rNote = cmd.execute()
		If Not rNote.Eof Then
			For each fld in rNote.Fields
				If Not IsNull(fld) Then myNote = Replace(myNote, "{" & fld.Name & "}", fld) Else myNote = Replace(myNote, "{" & fld.Name & "}", "")
			Next
		End If
	End If

	myNote = Replace(myNote, chr(13), "<br>")
	BuildMultiNote = myNote
End Function

Sub LoadMultiParams
	cmd.Parameters.Refresh()
	cmd("@LanID") = Session("LanID")
	cmd("@SlpCode") = Session("vendid")
	cmd("@dbName") = Session("olkdb")
	cmd("@branch") = Session("branch")
	cmd("@UserType") = "V"
	cmd("@LogNum") = Session("RetVal")
	cmd("@CardCode") = Session("UserName")
	cmd("@ItemCode") = ItemCode
	If Qty <> "" Then cmd("@Quantity") = CDbl(getNumericOut(Qty))
End Sub
 %>
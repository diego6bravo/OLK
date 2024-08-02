<% addLngPathStr = "cart/" %>
<!--#include file="lang/cartEditLine.asp" -->
<%
sql = 	"select T0.SDKID, " & _
		"Case When Exists(select 'A' from OLKCUFD where TableID = 'INV1' and Active = 'Y') Then 'Y' Else 'N' End EnableSDK, " & _
		"IsNull((select MaxDiscount from OLKAgentsAccess where SlpCode = " & Session("vendid") & "), 0) MaxLineDiscount " & _
		"from R3_ObsCommon..TCIF T0 cross join OLKCommon T1 " & _
		"where T0.CompanyDB = N'" & Session("olkdb") & "'"
set rs = conn.execute(sql)

If Not rs.Eof Then

SDKID = rs("SDKID")
EnableSDK = rs("EnableSDK") = "Y"
If Session("useraccess") = "U" Then
	MaxDiscount = rs("MaxLineDiscount")
Else
	MaxDiscount = 100
End If

sqlAddStr = ""
sqlAddStr2 = ""

sqlAddStr = sqlAddStr & ", IsNULL(T0.VatGroup,T1.VatGourpSa) VatGroup, T0.TaxCode "
sqlAddStr2 = "inner join R3_ObsCommon..TDOC T2 on T2.LogNum = T0.LogNum "

If myApp.SDKLineMemo Then
	If myApp.SVer < 8 Then
		sqlAddStr = sqlAddStr & ", " & SDKID & "LineMemo LineMemo "
	Else
		sqlAddStr = sqlAddStr & ", T3.LineText LineMemo "
		sqlAddStr2 = sqlAddStr2 & "left outer join R3_ObsCommon..DOC10 T3 on T3.LogNum = T0.LogNum and T3.LineType = 'T' and T3.AfterLine = T0.LineNum "
	End If
End If

			set rg = Server.CreateObject("ADODB.RecordSet")
			sql = "select T0.GroupID, IsNull(T1.AlterGroupName, T0.GroupName) GroupName " & _
					"from OLKCUFDGroups T0 " & _
					"left outer join OLKCUFDGroupsAlterNames T1 on T1.TableID = T0.TableID and T1.GroupID = T0.GroupID and T1.LanID = " & Session("LanID") & " " & _
					"where T0.TableID = 'OINV' and exists(select '' from CUFD X0 left outer join OLKCUFD X1 on X1.TableID = X0.TableID and X1.FieldID = X0.FieldID where X0.TableID = T0.TableID and IsNull(X1.GroupID, -1) = T0.GroupID and AType in ('" & userType & "', 'T') and OP in ('T','P') and Active = 'Y') " & _
					"order by T0.[Order] "
			set rg = conn.execute(sql)

set rSdk = Server.CreateObject("ADODB.RecordSet")
sql = "select IsNull(T1.GroupID, -1) GroupID, T0.FieldID, '" & SDKID & "'+AliasID InsertID, AliasID, IsNull(alterDescr, Descr) Descr, TypeID, SizeID, EditType, Dflt, NotNull, IsNull(T1.Pos, 'D') Pos, RTable, " & _
				  "Case When Exists(select 'A' from UFD1 where TableId = T0.TableId and FieldId = T0.FieldId) " & _
				  "Then 'Y' Else 'N' End As DropDown, NullField, Query " & _
				  "from cufd T0 " & _
				  "left outer join OLKCUFD T1 on T1.TableID = T0.TableID and T1.FieldID = T0.FieldID " & _
				  "left outer join OLKCUFDAlterNames T2 on T2.TableID = T0.TableID and T2.FieldID = T0.FieldID and T2.LanID = " & Session("LanID") & " " & _
				  "where T0.TableId = 'INV1' and AType in ('" & userType & "', 'T') and OP in ('T','P') and Active = 'Y' " & _
			 	  "order by IsNull(T1.GroupID, -1), IsNull(T1.Pos, 'D'), IsNull(T1.[Order], 32727) "
				  
If myApp.SVer < 8 Then sql = sql & " and T0.AliasID <> 'LineMemo'"
rSdk.open sql, conn, 3, 1

do while not rSdk.eof
	sqlAddStr = sqlAddStr & ", T0." & rSdk("InsertID")
rSdk.movenext
loop
If rSdk.recordcount >0 Then rSdk.movefirst

sql = 	"select T0.ItemCode, T1.ItemName, T0.WhsCode, T0.UnitPrice, T0.Price, Case T4.SaleType When 1 Then T0.Quantity When 2 Then T0.Quantity When 3 Then T0.Quantity/SalPackUn End Quantity, " & _
"T0.DiscPrcnt, T4.SaleType, T1.NumInSale, T1.SalPackUn " & sqlAddStr & " " & _
		"from R3_ObsCommon..DOC1 T0 " & _
		"inner join OITM T1 on T1.ItemCode = T0.ItemCode collate database_default " & sqlAddStr2 & _
		"INNER JOIN OlkSalesLines T4 on T4.LogNum = T0.Lognum and T4.LineNum = T0.LineNum " & _
		"where T0.LogNum = " & Session("RetVal") & " and T0.LineNum = " & Request("LineNum")
set rs = conn.execute(sql)

TaxGroup = ""
VatGroup = ""

If Request("err") = "" then
	If myApp.SDKLineMemo then LineMemo = RS("LineMemo")
	WhsCode = RS("WhsCode")
Else
	If myApp.SDKLineMemo then LineMemo = Request("LineMemo")
	WhsCode = Request("WhsCode")
End If

If Request("cmd") <> "e" Then
	VatGroup = rs("VatGroup")
	TaxCode = rs("TaxCode")
Else
	VatGroup = Request("VatGroup")
	TaxCode = Request("TaxCode")
End If

If myApp.UnEmbPriceSet and rs("SaleType") = 3 Then 
	Price = Replace(FormatNumber(CDbl(rs("Price"))*CDbl(rs("SalPackUn")),myApp.PriceDec),GetFormatSep(),"")
Else 
	Price = Replace(FormatNumber(CDbl(rs("Price")),myApp.PriceDec),GetFormatSep(),"")
End If

If myApp.UnEmbPriceSet and rs("SaleType") = 3 or rs("SaleType") <> 3 Then
	LineTotal = Replace(FormatNumber(CDbl(Price)*CDbl(rs("Quantity")), myApp.SumDec),GetFormatSep(),"")
Else
	LineTotal = Replace(FormatNumber(CDbl(Price)*CDbl(rs("Quantity"))*CDbl(rs("SalPackUn")), myApp.SumDec),GetFormatSep(),"")
End If
 %>
<script language="javascript">
var SaleType = <%=rs("SaleType")%>;
var NumInSale = <%=rs("NumInSale")%>;
var SalPackUn = <%=rs("SalPackUn")%>;

function updateNote() 
{
	document.frmEditLine.LineMemo.value = document.frmEditLine.NoteVar.options[document.frmEditLine.NoteVar.selectedIndex].value
}
function getCal(AliasID)
{
	document.frmEditLine.action = 'operaciones.asp';
	document.frmEditLine.editVar.value = AliasID;
	document.frmEditLine.cmd.value = 'UDFCal';
	document.frmEditLine.submit();
}
function getVal(AliasID)
{
	document.frmEditLine.action = 'operaciones.asp';
	document.frmEditLine.editVar.value = AliasID;
	document.frmEditLine.cmd.value = 'UDFQry';
	document.frmEditLine.submit();
}
<% If Request("err") = "Y" Then %>
alert("<%=getcartEditLineLngStr("LtxtValItmQty")%>")
<% end if %>
</script>
<form method="post" action="cart/cartEditLineSubmit.asp" name="frmEditLine">
<div align="center">
	<center>
		<table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
			<tr>
				<td bgcolor="#9BC4FF">
				<table border="0" cellpadding="0" bordercolor="#111111" width="100%">
					<tr>
						<td colspan="2">
						<p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
						<b><font face="Verdana" size="1"><%=Replace(Replace(getcartEditLineLngStr("LtxtShopCart"), "{0}", Session("RetVal")), "{1}", CInt(Request("LineNum"))+1)%></font></b>
						</td>
					</tr>
					<tr>
						<td width="33%" bgcolor="#7DB1FF"><b><font size="1" face="Verdana"><%=getcartEditLineLngStr("DtxtItem")%></font></b></td>
						<td><font size="1" face="Verdana"><%=rs("ItemCode")%></font></td>
					</tr>
					<tr>
						<td width="33%" bgcolor="#7DB1FF" valign="top"><b><font size="1" face="Verdana"><%=getcartEditLineLngStr("DtxtDescription")%></font></b></td>
						<td><font size="1" face="Verdana"><%=rs("ItemName")%></font></td>
					</tr>
					<tr>
						<td width="33%" bgcolor="#7DB1FF" valign="top"><b><font size="1" face="Verdana">
						<%=getcartEditLineLngStr("DtxtQty")%></font></b></td>
						<td align="right"><font size="1" face="Verdana"><%=Replace(FormatNumber(CDbl(rs("Quantity")), myApp.QtyDec), GetFormatSep(), "")%><input type="hidden" name="Quantity" value="<%=rs("Quantity")%>"></font></td>
					</tr>
					<tr>
						<td width="33%" bgcolor="#7DB1FF" valign="top"><b><font size="1" face="Verdana">
						<%=getcartEditLineLngStr("DtxtPrice")%></font></b></td>
						<td><input type="text" name="Price" <% If Not myAut.HasAuthorization(68) Then %> readonly <% End If %>  class="input"  value="<%=Price%>" style="text-align: right;width: 100%; font-family: Verdana; font-size: 10px" onfocus="this.select();" onmouseup="event.preventDefault()" onchange="chkNum(this, document.frmEditLine.oldPrice, <%=myApp.PriceDec%>);"><input type="hidden" name="oldPrice" value="<%=Price%>"></td>
					</tr>
					<% If myApp.ShowLineDiscount Then %>
					<tr>
						<td width="33%" bgcolor="#7DB1FF" valign="top"><b><font size="1" face="Verdana">
						<%=getcartEditLineLngStr("DtxtDiscount")%></font></b></td>
						<td><input type="text" name="Discount" <% If Not myAut.HasAuthorization(68) Then %> readonly <% End If %>  value="<%=FormatNumber(CDbl(rs("DiscPrcnt")), myApp.PercentDec)%>" style="text-align: right; width: 100%; font-family: Verdana; font-size: 10px" onfocus="this.select();" onmouseup="event.preventDefault()" onchange="chkNum(this, document.frmEditLine.oldDiscount, <%=myApp.PercentDec%>);"></td>
					</tr>
					<% Else %>
					<input type="hidden" name="Discount" value="<%=FormatNumber(CDbl(rs("DiscPrcnt")), myApp.PercentDec)%>">
					<% End If %>
					<input type="hidden" name="oldDiscount" value="<%=rs("DiscPrcnt")%>">
					<tr>
						<td width="33%" bgcolor="#7DB1FF" valign="top"><b><font size="1" face="Verdana">
						<%=getcartEditLineLngStr("DtxtTotal")%></font></b></td>
						<td><input type="text" readonly style="text-align: right;" class="input" name="LineTotal" value="<%=LineTotal%>" style="width: 100%; font-family: Verdana; font-size: 10px"></td>
					</tr>
					<tr>
						<td width="33%" bgcolor="#7DB1FF" valign="top"><b><font size="1" face="Verdana"><%=getcartEditLineLngStr("DtxtWarehouse")%></font></b></td>
						<td>
						<% If myAut.HasAuthorization(99) Then %><select name="WhsCode" size="1"><%
						set rw = Server.CreateObject("ADODB.recordset")
						set cmd = Server.CreateObject("ADODB.Command")
						cmd.ActiveConnection = connCommon
						cmd.CommandType = &H0004
						cmd.CommandText = "DBOLKGetWarehouses" & Session("ID")
						cmd.Parameters.Refresh()
						cmd("@LanID") = Session("LanID")
						rw.open cmd, , 3, 1
						do while not rw.eof %>
						<option value="<%=rw("WhsCode")%>" <% If rw(0) = WhsCode Then %>selected<% End If %>><%=rw("WhsName")%></option>
						<% rw.movenext
						loop %>
						</select><% Else
						set rw = Server.CreateObject("ADODB.recordset")
						sql = "select whscode, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OWHS', 'WhsName', WhsCode, WhsName) WhsName from owhs where WhsCode = N'" & saveHTMLDecode(WhsCode, False) & "'"
						set rw = conn.execute(sql) %><input type="hidden" name="WhsCode" value="<%=WhsCode%>"><font size="1" face="Verdana"><%=rw("WhsName")%></font><% End If %></td>
					</tr>
			      	<% If myApp.SVer >= 6 Then %>
			      	<% Select Case myApp.LawsSet 
			      		Case "PA", "AT", "AU", "BE", "CH", "CZ", "DE", "DK", "ES", "FI", "FR", "CN", "CY", "HU", "IT", "NL", "NO", "PL", "PT", "RU", "SE", "SK", "GB", "ZA" %>
					<tr>
						<td width="33%" bgcolor="#7DB1FF" valign="top"><b><font size="1" face="Verdana"><%=getcartEditLineLngStr("LtxtVatGrp")%></font></b></td>
						<td>
						<% If myAut.HasAuthorization(175) Then 
						sql = "select Code, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OVTG', 'Name', Code, Name) Name from OVTG where Category = 'O'"
				        set rw = conn.execute(sql) %>
				        <select size="1" name="VATGroup" style="width: 100%">
					    <% do while not rw.eof %>
						<option value="<%=myHTMLEncode(RW(0))%>" <% If Rw(0) = VATGroup Then %>selected<%end if %>><%=myHTMLEncode(RW(1))%></option>
						<% rw.movenext
						loop %>
				        </select>
				        <% Else %><font size="1" face="Verdana"><%
				        sql = "select OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OVTG', 'Name', Code, Name) Name from OVTG where Category = 'O' and Code = N'" & saveHTMLDecode(VATGroup, False) & "'"
				        set rw = conn.execute(sql)
				        Response.Write rw(0) %></font><input type="hidden" name="VATGroup" value="<%=myHTMLEncode(VATGroup)%>"><% End If %>
						</td>
					</tr>
					<% Case "MX", "CL", "CR", "GT", "US", "CA", "BR" %>
					<tr>
						<td width="33%" bgcolor="#7DB1FF" valign="top"><b><font size="1" face="Verdana"><%=getcartEditLineLngStr("LtxtTaxCode")%></font></b></td>
						<td><% If myAut.HasAuthorization(175) Then %>
						<select size="1" name="TaxCode" style="width: 100%">
						<% 
						sql = "select Code, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OSTC', 'Name', Code, Name) Name from OSTC where ValidForAR = 'Y'"
						set rw = conn.execute(sql)
						do while not rw.eof %>
						<option <% If TaxCode = rw("Code") Then %>selected<% End If %> value="<%=myHTMLEncode(rw("Code"))%>"><%=myHTMLEncode(rw("Code"))%> 
						- <%=myHTMLEncode(rw("Name"))%></option>
						<% rw.movenext
						loop %>
						</select>
				        <% Else %><font size="1" face="Verdana"><%
				        sql = "select OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OSTC', 'Name', Code, Name) Name from OSTC where ValidForAR = 'Y' and Code = N'" & saveHTMLDecode(TaxCode, False) & "'"
				        set rw = conn.execute(sql)
				        Response.Write rw(0) %></font><input type="hidden" name="TaxCode" value="<%=myHTMLEncode(TaxCode)%>"><% End If %>
						</td>
					</tr>
					<% End Select %>
					<% End If %>
					<tr>
						<td width="33%" bgcolor="#7DB1FF" valign="top"><b><font size="1" face="Verdana"><%=getcartEditLineLngStr("DtxtNote")%></font></b></td>
						<td>
						<select <% If Not myApp.SDKLineMemo Then %>disabled<% end if %> size="1" name="NoteVar" style="width: 100%" onchange="updateNote()">
						<option><%=getcartEditLineLngStr("LtxtPredNotes")%></option>
						<% set rw = conn.execute("select NoteName, Note from OLKNotePerLine")
						do while not rw.eof %>
						<option value="<%=myHTMLEncode(RW("Note"))%>"><%=myHTMLEncode(RW("NoteName"))%></option>
						<% rw.movenext
						loop %>
						</select></td>
					</tr>
					<tr>
						<td bgcolor="#7DB1FF" width="33%" valign="top"><b><font size="1" face="Verdana"><%=getcartEditLineLngStr("DtxtNote")%></font></b></td>
						<td><textarea <% If Not myApp.SDKLineMemo Then %>disabled<% end if %> rows="5" name="LineMemo" id="LineMemo" cols="14" onkeydown="return chkMax(event, this, 254);" style="width: 100%; "><%=myHTMLEncode(LineMemo)%><% If Not myApp.SDKLineMemo Then %><%=getcartEditLineLngStr("LtxtDisNotes")%><% end if %></textarea></td>
					</tr>
					<% do while not rg.eof %>
					<tr>
						<td bgcolor="#7DB1FF" colspan="2" valign="top"><b><font size="1" face="Verdana"><% Select Case CInt(rg("GroupID"))
						Case -1 %><%=getcartEditLineLngStr("DtxtUDF")%><%
						Case Else 
							Response.Write rg("GroupName")
						End Select %></font></b></td>
					</tr>
					<% rSdk.Filter = "GroupID = " & rg("GroupID")
					
					do while not rSdk.eof 
                    AliasID = rSdk("InsertID")
                    If Request.Form.Count = 0 Then
	                    fldVal = rs(AliasID)
	                    If rSdk("TypeID") = "D" Then fldVal = FormatDate(fldVal, False)
	                Else
	                	fldVal = Request("U_" & rSdk("AliasID"))
	                End If %>
            <tr>
              <td width="33%" bgcolor="#7DB1FF"><b>
                      <font size="1" face="Verdana">&nbsp;<%=rSdk("Descr")%><% If rSdk("NullField") = "Y" Then %><font color="red">*</font><% End If %></font></b></td>
              <td width="67%">
        		<% If rSdk("DropDown") = "Y" or Not IsNull(rSdk("RTable")) then 
        		If rSdk("DropDown") = "Y" Then
	        		sql = "select FldValue, IsNull(AlterDescr, Descr) Descr " & _
									"from UFD1 T0 " & _
									"left outer join OLKUFD1AlterNames T1 on T1.TableID = T0.TableID and T1.FieldID = T0.FieldID and T1.IndexID = T0.IndexID and T1.LanID = " & Session("LanID") & " " & _
									"where T0.tableid = 'INV1' and T0.FieldId = " & rcOpt("FieldId")
				Else
					sql = "select Code FldValue, Name Descr from [@" & rSdk("RTable") & "] order by 2"
				End If
				set rctn = conn.execute(sql) %>
				<font color="#4783C5">
				<select size="1" name="U_<%=rSdk("AliasID")%>" class="input" style="font-size:10px; width:100%; font-family:Verdana">
				<option></option>
				<% do while not rctn.eof %>
				<option value="<%=rctn("FldValue")%>" <% If fldVal = rctn("FldValue") Then %>selected<% ElseIf rctn("FldValue") = rSdk("Dflt") and IsNull(fldVal) Then %>selected<% End If %>><%=rctn("Descr")%></option>
				<% rctn.movenext
				loop
				rctn.close %></select></font><font size="1" color="#4783C5">
				<% Else %>
				<% If rSdk("Query") = "Y" or rSdk("TypeID") = "D" Then %><table width="100%" cellspacing="0" cellpadding="0"><tr><td width="16"><img border="0" src="<% If rSdk("Query") = "Y" Then %>../images/<%=Session("rtl")%>flechaselec2.gif<% Else %>images/cal.gif<% End If %>" <% If rSdk("TypeID") = "D" Then %>onclick="javascript:getCal('<%=rSdk("AliasID")%>')"<% End If %> <% If rSdk("Query") = "Y" Then %>onclick="javascript:getVal('<%=rSdk("AliasID")%>')"<% End If %>></td><td><% End If %>
				<input <% If rSdk("Query") = "Y" or rSdk("TypeID") = "D" Then %>readonly<% End If %> type="text" name="U_<%=rSdk("AliasID")%>" size="<% If rSdk("TypeID") = "A" Then %>43<% Else %>12<% End If %>" class="input" value="<% If fldVal <> "" Then %><%=fldVal%><% Else %><%=rSdk("Dflt")%><% End If %>" onclick="this.selectionStart=0;this.selectionEnd=this.value.length;" onchange="chkThis(this, '<%=rSdk("TypeID")%>', '<%=rSdk("EditType")%>', <%=rSdk("SizeID")%>)" style="width: 100%; font-family: Verdana; font-size: 10px">
				<% If rSdk("Query") = "Y" or rSdk("TypeID") = "D" Then %></td><td width="16"><img border="0" src="../images/remove.gif" width="16" height="16" onclick="javascript:document.frmEditLine.U_<%=rSdk("AliasID")%>.value = ''"></td></tr></table><% End If %><% End If %></td>
            </tr>
                    <% 
                    rSdk.movenext
                    loop 
                    rg.movenext
                    loop %>
				</table>
				</td>
			</tr>
	        <tr>
	          <td width="100%" bgcolor="#9BC4FF">
	          <div align="center">
	            <center>
	            <table border="0" cellpadding="0" cellspacing="1" style="border-collapse: collapse" bordercolor="#111111" width="100%">
	              <tr>
	                <td>
	                <p align="center">
	                <input border="0" src="images/save_icon.gif" name="I1" type="image"></td>
	                <td>
	                <p align="center"><a href="operaciones.asp?cmd=cart"><img border="0" src="images/x_icon.gif"></a></td>
	              </tr>
	            </table>
	            </center>
	          </div>
	          </td>
	        </tr>
		</table>
	</center>
</div>
<input type="hidden" name="LineNum" value="<%=Request("LineNum")%>">
<input type="hidden" name="editVar" value="">
<input type="hidden" name="returnCmd" value="cartEditLine">
<input type="hidden" name="cmd" value="cartEditLine">
<input type="hidden" name="ManPrc" value="N">
<input type="hidden" name="UnitPrice" value="<%=rs("UnitPrice")%>">
<input type="hidden" name="SaleType" value="<%=rs("SaleType")%>">
<input type="hidden" name="SalPackUn" value="<%=rs("SalPackUn")%>">
</form>
<script language="javascript">
var changePrice = true;
var formatDec = '<%=GetFormatDec()%>';
function IsNumeric(sText)
{
   var ValidChars = '0123456789.-';
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
function FormatNumber(expr, decplaces) 
{
	return formatNumberDec(expr, decplaces, false);
}
function chkNum(Fld, Old, Dec)
{
	MaxDiscount = '<%=FormatNumber(MaxDiscount, myApp.PercentDec)%>';
	var fldVal = Fld.value.replace(formatDec, '.');
	
	if (!IsNumeric(fldVal))
	{
		alert('<%=getcartEditLineLngStr("DtxtValNumVal")%>');
		Fld.value = Old.value;
		return;
	}
	else if (Fld.name == 'Price' && parseFloat(fldVal) < 0)
	{
		alert('<%=getcartEditLineLngStr("DtxtValNumMinVal")%>'.replace('{0}', '0'));
		Fld.value = FormatNumber(0, Dec);
	}
	Fld.value = FormatNumber(fldVal, Dec);
	
	switch (Fld.name)
	{
		case 'Discount':
			if (parseFloat(fldVal) > parseFloat(MaxDiscount.replace(formatDec, '.')))
			{
				alert('<%=getcartEditLineLngStr("LtxtMaxDiscount")%>'.replace('{0}', MaxDiscount));
				Fld.value = MaxDiscount;
				changePrice = true;
			}
			if (changePrice) setDiscPrice();
			break;
		case 'Price':
			setDiscValue();
			break;
	}
	
	Old.value = Fld.value;
	
	setLineTotal();
}

function setDiscValue()
{
	UnitPrice = parseFloat(document.frmEditLine.UnitPrice.value.replace(formatDec, '.'));
	Price = parseFloat(document.frmEditLine.Price.value.replace(formatDec, '.'));
	
	if (UnitPrice != 0)
	{
		SaleType = <%=rs("SaleType")%>;
		if (SaleType == 3)
		{
			UnitPrice = UnitPrice<% If Not myApp.UnEmbPriceSet Then %>*SalPackUn<% End If %>;
		}
		document.frmEditLine.Discount.value = FormatNumber(100-(Price*100)/UnitPrice, <%=myApp.PercentDec%>);
	}
	else
	{
		document.frmEditLine.Discount.value = FormatNumber(0, <%=myApp.PercentDec%>);
	}
	changePrice = false;
	chkNum(document.frmEditLine.Discount, document.frmEditLine.oldDiscount, <%=myApp.PercentDec%>);
	changePrice = true;
}

function setDiscPrice()
{
	var disc = parseFloat(document.frmEditLine.Discount.value);
	
	document.frmEditLine.ManPrc.value = 'Y';
	
	UnitPrice = parseFloat(document.frmEditLine.UnitPrice.value);
	if (UnitPrice != 0)
	{
		SaleType = <%=rs("SaleType")%>;
		if (SaleType == 3)
		{
			UnitPrice = UnitPrice*NumInSale<% If Not myApp.UnEmbPriceSet Then %>*SalPackUn<% End If %>;
		}
		document.frmEditLine.Price.value = FormatNumber(UnitPrice-((disc*UnitPrice)/100), <%=myApp.PriceDec%>);
	}
	else
	{
		document.frmEditLine.Price.value = FormatNumber(0, <%=myApp.PriceDec%>);
	}
	document.frmEditLine.oldPrice.value = document.frmEditLine.Price.value;
}

function setLineTotal()
{
	document.frmEditLine.LineTotal.value = FormatNumber(parseFloat(document.frmEditLine.Quantity.value)*parseFloat(document.frmEditLine.Price.value), <%=myApp.SumDec%>);
}

function chkThis(Field, FType, EditType, FSize)
{
	switch (FType)
	{
		case 'A':
			if (Field.value.length > FSize)
			{
				alert('<%=getcartEditLineLngStr("DtxtValFldMaxChar")%>'.replace('{0}', FSize));
				Field.value = Left(Field.value, FSize)
			}
			break;
		case 'N':
			if (EditType == ' ')
			{
				if (Field.value != '')
				{
					if (!MyIsNumeric(Field.value))
					{
						Field.value = '';
						alert('<%=getcartEditLineLngStr("DtxtValNumVal")%>');
					}
					else if (parseFloat(Field.value) > 2147483647)
					{
						alert('<%=getcartEditLineLngStr("DtxtValNumMaxVal")%>'.replace('{0}', '2147483647'));
						Field.value = 2147483647 ;
					}
					else if (parseFloat(Field.value) - parseInt(Field.value) != 0)
					{
						Field.value = '';
						alert('<%=getcartEditLineLngStr("DtxtValNumValWhole")%>');
					}
				}
			}
			break;
		case 'B':
			if (Field.value != '')
			{
				if (!MyIsNumeric(Field.value))
				{
					Field.value = '';
					alert('<%=getcartEditLineLngStr("DtxtValNumVal")%>');
				}
				else
				{
					if (parseFloat(Field.value) > 1000000000000)
					{
						Field.value = 999999999999;
					}
					else if (parseFloat(Field.value) < -1000000000000)
					{
						Field.value = -999999999999;
					}
					switch (EditType)
					{
						case 'R':
							Field.value = OLKFormatNumber(parseFloat(Field.value), <%=myApp.RateDec%>);
							break;
						case 'S':
							Field.value = OLKFormatNumber(parseFloat(Field.value), <%=myApp.SumDec%>);
							break;
						case 'P':
							Field.value = OLKFormatNumber(parseFloat(Field.value), <%=myApp.PriceDec%>);
							break;
						case 'Q':
							Field.value = OLKFormatNumber(parseFloat(Field.value), <%=myApp.QtyDec%>);
							break;
						case '%':
							Field.value = OLKFormatNumber(parseFloat(Field.value), <%=myApp.PercentDec%>);
							break;
						case 'M':
							Field.value = OLKFormatNumber(parseFloat(Field.value), <%=myApp.MeasureDec%>);
							break;
					}
				}
			}
			break;
	}
}
</script>
<% Else
sql = "select case when exists(select '' from R3_ObsCommon..TCIF where CompanyDB = db_name()) Then 0 Else 1 End ErrID"
set rs = conn.execute(sql)
ErrID = rs(0) '0 = Unknown, 1 = Database missing in observer
 %>
<%=getcartEditLineLngStr("DtxtDataError")%>&nbsp;<% Select Case ErrID 
Case 0 %><%=getcartEditLineLngStr("DtxtUnknown")%>
<% Case 1 %><%=getcartEditLineLngStr("LtxtDBObsErr")%>
<% End Select %><% End If %>
<head>
<style type="text/css">
.style1 {
	text-align: right;
}
</style>
</head>

<% addLngPathStr = "cart/" %>
<!--#include file="lang/cart_extra.asp" -->
<%
varx = 0
set rc = server.createobject("ADODB.RecordSet")
set rd = server.createobject("ADODB.RecordSet")
set rctn = server.createobject("ADODB.RecordSet")

sql = "select cntctcode, comments, T0.CardCode, T0.CardName, E_Mail, object, IsNull(T0.ReserveInvoice, 'N') ReserveInvoice, T1.currency, NumAtCard, T0.slpcode, IsNull(T0.GroupNum, -1) GroupNum, T2.Object, " & _
	  "DocDate, DocDueDate, PartSupply, IsNull(T0.ShipToCode, T1.ShipToDef) ShipToCode, OLKCommon.dbo.DBOLKFormatAddress" & Session("ID") & "(T0.CardCode, 'S', IsNull(T0.ShipToCode, T1.ShipToDef), " & Session("LanID") & ") ShipAddress, " & _
	  "IsNull(T0.PayToCode, T1.BillToDef) PayToCode, OLKCommon.dbo.DBOLKFormatAddress" & Session("ID") & "(T0.CardCode, 'B', IsNull(T0.PayToCode, T1.BillToDef), " & Session("LanID") & ") PayAddress, T0.Project, " & _
	  "Case when T1.Currency <> N'" & myApp.MainCur & "' Then 'Y' Else 'N' End ShowCurncy, " & _
	  "Case When T1.Currency = '##' Then 'Y' Else 'N' End EnableMC, T0.DocCur " & _
	  "from R3_ObsCommon..tdoc T0 " & _
	  "inner join ocrd T1 on T1.cardcode = T0.cardcode collate database_default " & _
	  "inner join r3_obscommon..tlog T2 on T2.lognum = T0.lognum " & _
	  "where T0.lognum = " & Session("RetVal")

set rs = conn.execute(sql)
PartSupply = rs("PartSupply")
ShowCurrency = rs("ShowCurncy") = "Y"
EnableMC = rs("EnableMC")


			set rg = Server.CreateObject("ADODB.RecordSet")
			sql = "select T0.GroupID, IsNull(T1.AlterGroupName, T0.GroupName) GroupName " & _
					"from OLKCUFDGroups T0 " & _
					"left outer join OLKCUFDGroupsAlterNames T1 on T1.TableID = T0.TableID and T1.GroupID = T0.GroupID and T1.LanID = " & Session("LanID") & " " & _
					"where T0.TableID = 'OINV' and exists(select '' from CUFD X0 left outer join OLKCUFD X1 on X1.TableID = X0.TableID and X1.FieldID = X0.FieldID where X0.TableID = T0.TableID and IsNull(X1.GroupID, -1) = T0.GroupID and AType in ('" & userType & "', 'T') and OP in ('T','P') and Active = 'Y') " & _
					"order by T0.[Order] "
			set rg = conn.execute(sql)

			set rcOpt = Server.CreateObject("ADODB.RecordSet")
			sql = "select IsNull(T1.GroupID, -1) GroupID, T0.FieldID, AliasID, IsNull(alterDescr, Descr) Descr, TypeID, SizeID, Dflt, NotNull, IsNull(T1.Pos, 'D') Pos, RTable, " & _
				  "Case When Exists(select 'A' from UFD1 where TableId = T0.TableId and FieldId = T0.FieldId) " & _
				  "Then 'Y' Else 'N' End As DropDown, NullField, Query, " & _
				  "(select SDKID collate database_default from r3_obscommon..tcif where companydb = '" & Session("OlkDB") & "')++AliasID As InsertID " & _
				  "from cufd T0 " & _
				  "left outer join OLKCUFD T1 on T1.TableID = T0.TableID and T1.FieldID = T0.FieldID " & _
				  "left outer join OLKCUFDAlterNames T2 on T2.TableID = T0.TableID and T2.FieldID = T0.FieldID and T2.LanID = " & Session("LanID") & " " & _
				  "where T0.TableId = 'OINV' and AType in ('" & userType & "', 'T') and OP in ('T','P') and Active = 'Y' " & _
			 	  "order by IsNull(T1.GroupID, -1), IsNull(T1.Pos, 'D'), IsNull(T1.[Order], 32727) "

			rcOpt.open sql, conn, 3, 1

			sql = "select AliasID, NullField, Descr, TypeID, SizeID " & _
				  "from cufd T0 " & _
				  "left outer join OLKCUFD T1 on T1.TableID = T0.TableID and T1.FieldID = T0.FieldID " & _
				  "where T0.TableId = 'OINV' and AType in ('" & userType & "', 'T') and OP in ('T','P') and Active = 'Y' and (NullField = 'Y' or TypeID in ('N', 'B')) " & _
				  "Order By Pos Desc"
			rctn.open sql, conn, 3, 1
			
			If rctn.recordcount > 0 Then chkOpt = True
			
			If rcOpt.RecordCount > 0 Then
				set rcOptVals = Server.CreateObject("ADODB.RecordSet")
				sql = "select (select SDKID collate database_default from r3_obscommon..tcif where companydb = '" & Session("OlkDB") & "')++AliasID As InsertID, TypeID " & _
					  "from cufd T0 " & _
					  "left outer join OLKCUFD T1 on T1.TableID = T0.TableID and T1.FieldID = T0.FieldID " & _
					  "where T0.TableId = 'OINV' and AType in ('" & userType & "','T') and OP in ('T','P')  and Active = 'Y'"
				rcOptVals.open sql, conn, 3, 1
				sql = "select "
				do while not rcOptVals.eof
					If rcOptVals.bookmark <> 1 Then sql = sql & ", "
					sql = sql & rcOptVals("InsertID")
				rcOptVals.movenext
				loop
				sql = sql & " from r3_obscommon..tdoc where lognum = " & Session("RetVal")
				set rcOptVals = conn.execute(sql)
			End If

If Request.Form.Count = 0 Then
	DocDate 	= FormatDate(rs("DocDate"), False)
	DocDueDate 	= FormatDate(rs("DocDueDate"), False)
	CardName	= rs("CardName")
	If rs("CntctCode") <> "" Then
		CntctCode	= CInt(rs("CntctCode"))
	Else
		CntctCode = -1
	End If
	NumAtCard	= rs("NumAtCard")
	SlpCode		= CInt(rs("SlpCode"))
	Project		= rs("Project")
	GroupNum	= CInt(RS("GroupNum"))
	Comments	= RS("Comments")
	ObjType 	= rs("Object")
	ReserveInvoice = rs("ReserveInvoice")
	PayToCode = rs("PayToCode")
	PayAddress = RS("PayAddress")
	ShipToCode = rs("ShipToCode")
	ShipAddress = RS("ShipAddress")
	DocCur = rs("DocCur")
Else
	DocDate 	= Request("DocDate")
	DocDueDate 	= Request("DocDueDate")
	CardName	= Request("CardName")
	If Request("CntctCode") <> "" Then
		CntctCode	= CInt(Request("CntctCode"))
	Else
		CntctCode = -1
	End If
	NumAtCard	= Request("NumAtCard")
	SlpCode		= CInt(Request("SlpCode"))
	Project		= rs("Project")
	GroupNum	= CInt(Request("GroupNum"))
	Comments	= Request("Comments")
	ObjType		= CInt(Request("ObjType"))
	DocCur 		= Request("DocCur")
	
	PayToCode = Request("PayToCode")
	ShipToCode = Request("ShipToCode")
	sql = "select OLKCommon.dbo.DBOLKFormatAddress" & Session("ID") & "(N'" & saveHTMLDecode(Session("UserName"), False) & "', 'B', N'" & saveHTMLDecode(PayToCode, False) & "', " & Session("LanID") & ") [PayAddress], " & _
			"OLKCommon.dbo.DBOLKFormatAddress" & Session("ID") & "(N'" & saveHTMLDecode(Session("UserName"), False) & "', 'S', N'" & saveHTMLDecode(ShipToCode, False) & "', " & Session("LanID") & ") [ShipAddress] "
	set rd = conn.execute(sql)
	PayAddress = rd("PayAddress")
	ShipAddress = rd("ShipAddress")
	
	If ObjType = -13 Then
		ObjType = 13
		ReserveInvoice = "Y"
	Else
		ReserveInvoice = "N"
	End If
	
	If Request.Form("changeObj") = "Y" Then
		sql = "declare @obj int set @obj = " & ObjType & " " & _  
				"declare @GroupNum int set @GroupNum = " & GroupNum & " " & _  
				"declare @DocDate datetime set @DocDate = Convert(datetime,'" & SaveSqlDate(Request("DocDate")) & "',120) " & _
				"select Case @obj  " & _  
				"	When 13 Then " & _  
				"	DateAdd(day,ExtraDays, " & _  
				"		DateAdd(month,ExtraMonth, " & _  
				"			Case PayDuMonth  " & _  
				"				When 'N' Then @DocDate  " & _  
				"				When 'Y' Then DateAdd(day,1-day(DateAdd(month,1,@DocDate)),DateAdd(month,1,@DocDate)) " & _  
				"				When 'H' Then DateAdd(day,15-day(DateAdd(month,1,@DocDate)),DateAdd(month,1,@DocDate)) " & _  
				" 				When 'E' Then DateAdd(day,1-day(DateAdd(month,1,@DocDate)),DateAdd(month,1,@DocDate))-1  " & _  
				"			End " & _  
				"		) " & _  
				"	) " & _  
				"	When 15 Then @DocDate " & _  
				"	When 17 Then null " & _  
				"	When 23 Then DateAdd(month, 1, @DocDate) " & _  
				"End DocDueDate " & _  
				"from OCTG " & _  
				"where GroupNum = @GroupNum " 
		set rd = conn.execute(sql)
		DocDueDate = FormatDate(rd(0), False)
		rd.close
	End If
End If
			
%>
<script language="javascript">
function checkPList(i)
{
	var plist = parseInt(document.frmCart.NewPriceList.value);
	var value = parseInt(ListNum[i]);
	
	if (plist != value)
	{
		if (confirm("<%=getcart_extraLngStr("LtxtApplyPListByPTerm")%>"))
		{
			document.frmCart.NewPriceList.value = value;
			document.frmCart.ChangePList.value = 'Y';
		}
	}
}
function ValidateForm() {
	<% do while not rctn.eof
	If rctn("NullField") = "Y" Then %>
	if (document.frmCart.U_<%=rctn("AliasID")%>.value == '') {
		alert('<%=getcart_extraLngStr("LtxtValFld")%>'.replace('{0}', '<%=Replace(rctn("Descr"), "'", "\'")%>'));
		document.frmCart.U_<%=rctn("AliasID")%>.focus
		return; }
	<% End If
	If rctn("TypeID") = "B" or rctn("TypeID") = "N" Then %>
	if (document.frmCart.U_<%=rctn("AliasID")%>.value != '') 
	{
		if (!IsNumeric(document.frmCart.U_<%=rctn("AliasID")%>.value)) 
		{
			alert('<%=getcart_extraLngStr("DtxtValNumVal")%>');
			document.frmCart.U_<%=rctn("AliasID")%>.focus
			return false; 
		}
	}
	<% End If
	rctn.movenext
	loop %>
	document.frmCart.cmd.value = 'cartExtra';
	document.frmCart.action = 'cart/cartupdate2.asp';
	document.frmCart.submit();
}

function IsNumeric(sText)
{
   var ValidChars = "0123456789<%=GetFormatDec()%>";
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

   
function getCal(AliasID, System)
{
	document.frmCart.action = 'operaciones.asp';
	document.frmCart.editVar.value = AliasID;
	document.frmCart.cmd.value = 'UDFCal';
	document.frmCart.System.value = System;
	document.frmCart.submit();
}
function getVal(AliasID)
{
	document.frmCart.action = 'operaciones.asp';
	document.frmCart.editVar.value = AliasID;
	document.frmCart.cmd.value = 'UDFQry';
	document.frmCart.submit();
}
function doObj()
{
	document.frmCart.changeObj.value = "Y";
	document.frmCart.action = 'operaciones.asp';
	document.frmCart.cmd.value = 'cartopt';
	document.frmCart.submit();
}
function doReload()
{
	document.frmCart.action = 'operaciones.asp';
	document.frmCart.cmd.value = 'cartopt';
	document.frmCart.submit();
}
</script>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
    <tr>
      <td bgcolor="#9BC4FF" class="style1">
      <form method="POST" action="operaciones.asp" name="frmCart">
      <input type="hidden" name="changeObj" value="N">
      <input type="hidden" name="System" value="">
        <table border="0" cellpadding="0" bordercolor="#111111" width="100%">
        <tr>
          <td width="100%">
          <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><%=getcart_extraLngStr("LtxtShopCartDet")%></font></b></td>
        </tr>
        <tr>
          <td width="100%">
          <table border="0" cellpadding="0" bordercolor="#111111" width="100%">
		<tr>
              <td><b><font size="1" face="Verdana"><nobr><%=getcart_extraLngStr("DtxtLogNum")%></nobr></font></b></td>
              <td colspan="2"><font size="1" face="Verdana"><%=Session("RetVal")%></font></td>
            </tr>
		<tr>
              <td colspan="3">
              <table cellpadding="0" border="0" width="100%">
              <% If myApp.EnableOQUT Then %>
              <tr>
				<td bgcolor="#7DB1FF"><input type="radio" name="ObjType" value="23" id="ObjType23"<% If ObjType = "23" Then Response.Write " checked"%> onclick="javascript:doObj();"><b><font size="1" face="Verdana"><label for="ObjType23"><%=txtQuote%></label></font></b></td>
			  </tr>
			  <% End If %>
			  <% If myApp.EnableORDR Then %>
			  <tr>
          		<td bgcolor="#7DB1FF"><input type="radio" value="17" name="ObjType" id="ObjType17"<% If ObjType = "17" Then Response.Write " checked"%> onclick="javascript:doObj();"><b><font size="1" face="Verdana"><label for="ObjType17"><%=txtOrdr%></label></font></b></td>
              </tr>
              <% End If %>
              <% If myApp.EnableOINV Then %>
			  <tr>
          		<td bgcolor="#7DB1FF"><input type="radio" value="13" name="ObjType" id="ObjType13"<% If ObjType = "13" and ReserveInvoice = "N" Then Response.Write " checked"%> onclick="javascript:doObj();"><b><font size="1" face="Verdana"><label for="ObjType13"><%=txtInv%></label></font></b></td>
              </tr>
              <% End If %>
              <% If myApp.EnableOINVRes Then %>
			  <tr>
          		<td bgcolor="#7DB1FF"><input type="radio" value="-13" name="ObjType" id="ObjType_13"<% If ObjType = "13" and ReserveInvoice = "Y" Then Response.Write " checked"%> onclick="javascript:doObj();"><b><font size="1" face="Verdana"><label for="ObjType_13"><%=txtInvRes%></label></font></b></td>
              </tr>
              <% End If %>
              <% If myAut.HasAuthorization(104) and ObjType = 17 Then %>
			  <tr>
			  	<td bgcolor="#7DB1FF"><input type="checkbox" name="PartSupply" id="PartSupply" <% If PartSupply = "Y" Then %>checked<% End IF %> value="Y"><b><font size="1" face="Verdana"><label for="PartSupply"><%=getcart_extraLngStr("LtxtPartSupply")%></label></font></b></td>
              </tr>
              <% Else %>
              <input type="hidden" name="PartSupply" value="<%=PartSupply%>">
              <% End If %>
              <input type="hidden" name="ReserveInvoice" value="<%=ReserveInvoice%>">
              </table>
              </td>
            </tr>
            <tr>
              <td width="33%" bgcolor="#7DB1FF"><b><font size="1" face="Verdana"><%=getcart_extraLngStr("DtxtDate")%></font></b></td>
              <td width="67%" colspan="2">
              <table cellpadding="0" cellspacing="0" width="100%" border="0">
              	<tr>
              		<td width="16"><a href="#" onclick="javascript:getCal('DocDate', 'Y')"><img border="0" src="images/cal.gif" border="0"></a></td>
              		<td><input type="text" name="DocDate" size="20" readonly style="font-family: Verdana; font-size: 10px; width: 100%;" value="<%=DocDate%>"></td>
              	</tr>
              </table>
            </tr>
            <tr>
              <td width="33%" bgcolor="#7DB1FF"><b><font size="1" face="Verdana"><% Select Case ObjType 
		                      	Case 13
		                      		txtDueDate = getcart_extraLngStr("LtxtPymntDue")
		                      	Case 17
		                      		txtDueDate = getcart_extraLngStr("LtxtDelDate")
		                      	Case 23
		                      		txtDueDate = getcart_extraLngStr("LtxtComDate")
		                      	End Select %>
								<%=txtDueDate%></font></b></td>
              <td width="67%" colspan="2">
              <table cellpadding="0" cellspacing="0" width="100%" border="0">
              	<tr>
              		<td width="16"><a href="#" onclick="javascript:getCal('DocDueDate', 'Y')"><img border="0" src="images/cal.gif" border="0"></a></td>
              		<td><input type="text" name="DocDueDate" size="20" readonly style="font-family: Verdana; font-size: 10px; width: 100%;" value="<%=DocDueDate%>"></td>
              	</tr>
              </table>
            </tr>
            <% If ShowCurrency Then %>
            <tr>
              <td width="33%" bgcolor="#7DB1FF"><b><font size="1" face="Verdana"><%=getcart_extraLngStr("LtxtCurr")%></font></b></td>
              <td width="67%" colspan="2">
              <% If EnableMC = "N" Then %><input type="text" size="3" class="InputDes" readonly name="DocCur" value="<%=myHTMLEncode(DocCur)%>"><% 
              Else
			     %><select size="1" name="DocCur" class="input" style="font-family: Verdana; font-size: 10px;" onchange="doReload();"><%
				set cmd = Server.CreateObject("ADODB.Command")
				cmd.ActiveConnection = connCommon
				cmd.CommandType = &H0004
				cmd.CommandText = "DBOLKGetCurrencies" & Session("ID")
				cmd.Parameters.Refresh()
				cmd("@LanID") = Session("LanID")
				set rctn = cmd.execute()
				do while not rctn.eof %>
				<option <% If DocCur = rctn(0) Then %>selected<% End If %> value="<%=myHTMLEncode(rctn(0))%>"><%=myHTMLEncode(rctn(1))%></option>
				<% rctn.movenext
				loop %>
				</select><% End If
				CurRate = 0
				If DocCur <> myApp.MainCur Then
					set cmd = Server.CreateObject("ADODB.Command")
					cmd.ActiveConnection = connCommon
					cmd.CommandType = &H0004
					cmd.CommandText = "DBOLKGetCartCurrRate" & Session("ID")
					cmd.Parameters.Refresh()
					cmd("@LanID") = Session("LanID")
					cmd("@LogNum") = Session("RetVal")
					If Request("DocCur") <> "" Then cmd("@DocCur") = Request("DocCur")
					set rctn = cmd.execute()
					CurRate = CDbl(rctn("Rate"))
				End If %><input type="text" size="12" class="InputDes" readonly name="DocCurRate" id="DocCurRate" style="font-family: Verdana; font-size: 10px; text-align:right;<% If DocCur = myApp.MainCur Then %>display: none;<% End If %>" value="<%=FormatNumber(CurRate, myApp.RateDec)%>">
				</tr>
            <% Else %>
			<input type="hidden" name="DocCur" value="<%=myHTMLEncode(DocCur)%>">
            <% End If %>
            <tr>
              <td width="33%" bgcolor="#7DB1FF">
              <table cellpadding="0" cellspacing="0" border="0" width="100%">
				<tr>
					<td><b><font size="1" face="Verdana"><%=getcart_extraLngStr("DtxtFor")%></font></b>
					</td>
					<td align="right" width="15"><a href="operaciones.asp?cmd=datos&card=<%=CleanItem(rs("CardCode"))%>">
					<img border="0" src="images/<%=Session("rtl")%>flechaselec.gif"></a></td>
				</tr>
			</table>
              </td>
              <td width="7%" style="width: 57%"><input type="text" name="CardName" size="16" style="font-family: Verdana; font-size: 10px; width: 100%" value="<%=Replace(CardName, """", "&quot;")%>" onclick="this.selectionStart=0;this.selectionEnd=this.value.length;"></td>
              <td width="10%">
               <% if rs("E_Mail") <> "" then %><a href="mailto:<%=rs("E_Mail")%>"><img border="0" src="cart/mail_icon.gif"></a><%end if%></td>
            </tr>
            <tr>
              <td width="33%" bgcolor="#7DB1FF" valign="top">
              <b><font size="1" face="Verdana"><%=getcart_extraLngStr("LtxtShipAdd")%></font></b></td>
              <td style="width: 57%">
						<table border="0" cellpadding="0" width="100%" cellspacing="0">
							<tr>
								<td>
								<select class="input" size="1" name="ShipToCode" style="font-size:10px; font-family:Verdana; width:100%" onchange="doReload();">
								<% 
								set cmd = Server.CreateObject("ADODB.Command")
								cmd.ActiveConnection = connCommon
								cmd.CommandType = &H0004
								cmd.CommandText = "DBOLKGetBPAdds" & Session("ID")
								cmd.Parameters.Refresh()
								cmd("@CardCode") = Session("UserName")
								cmd("@Type") = "S"
								set rctn = cmd.execute()
								do while not rctn.eof %>
								<option value="<%=myHTMLEncode(rctn(0))%>" <% If ShipToCode = rctn(0) Then %>selected<% End If %>><%=myHTMLEncode(rctn(0))%></option>
								<% rctn.movenext
								loop %>
								</select></td>
							</tr>
							<tr>
								<td class="CanastaTbl"><font size="1" face="Verdana"><%=ShipAddress%></font></td>
							</tr>
						</table>
			  </td>
              <td width="10%">
               &nbsp;</td>
            </tr>
            <tr>
              <td width="33%" bgcolor="#7DB1FF" valign="top">
              <b><font size="1" face="Verdana"><%=getcart_extraLngStr("LtxtPayAdd")%></font></b></td>
              <td style="width: 57%">
						<table border="0" cellpadding="0" width="100%" cellspacing="0">
							<tr>
								<td>
								<select class="input" size="1" name="PayToCode" style="font-size:10px; font-family:Verdana; width:100%" onchange="doReload();">
								<% 
								cmd("@Type") = "B"
								set rctn = cmd.execute()
								do while not rctn.eof %>
								<option value="<%=myHTMLEncode(rctn(0))%>" <% If PayToCode = rctn(0) Then %>selected<% End If %>><%=myHTMLEncode(rctn(0))%></option>
								<% rctn.movenext
								loop %>
								</select></td>
							</tr>
							<tr>
								<td class="CanastaTbl"><font size="1" face="Verdana"><%=PayAddress%></font></td>
							</tr>
						</table>
              </td>
              <td width="10%">
               &nbsp;</td>
            </tr>
            <tr>
              <td width="33%" bgcolor="#7DB1FF"><b><font size="1" face="Verdana"><%=getcart_extraLngStr("DtxtContact")%></font></b></td>
              <td colspan="2">
        <select size="1" name="CntctCode" style="font-size:10px; width:100%; font-family:Verdana">
        <%  
        set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKGetBPContacts" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@CardCode") = rs("CardCode")
        set rc = cmd.execute()
        do while not rc.eof %>
               <option <%If CntctCode = CInt(rc("cntctcode")) then response.write "selected"%> value="<%=rc("cntctcode")%>"><%=rc("Name")%></option>
               <%rc.movenext
               loop %>
               </select></td>
            </tr>
            <tr>
              <td width="33%" bgcolor="#7DB1FF"><b><font size="1" face="Verdana"><%=txtRef2%></font></b></td>
              <td width="67%" colspan="2">
        <input type="text" name="NumAtCard" size="20" style="font-family: Verdana; font-size: 10px; width: 100%;" value="<%=NumAtCard%>" onclick="this.selectionStart=0;this.selectionEnd=this.value.length;"></td>
            </tr>
            <tr>
              <td width="33%" bgcolor="#7DB1FF"><b><font size="1" face="Verdana"><%=getcart_extraLngStr("DtxtAgent")%></font></b></td>
              <td width="67%" colspan="2">
              <% If myAut.HasAuthorization(96) Then %>
		        <select size="1" name="SlpCode" style="font-size:10px; width:100%; font-family:Verdana">
		        <%  
				set cmd = Server.CreateObject("ADODB.Command")
				cmd.ActiveConnection = connCommon
				cmd.CommandType = &H0004
				cmd.CommandText = "DBOLKGetAgents" & Session("ID")
				cmd.Parameters.Refresh()
				cmd("@LanID") = Session("LanID")
				set rc = cmd.execute()
		
		        do while not rc.eof %>
               <option <% If SlpCode = CInt(rc("SlpCode")) then response.write "selected"%> value="<%=rc("SlpCode")%>"><%=myHTMLEncode(rc("SlpName"))%></option>
               <%rc.movenext
               loop %>
               </select><%
               Else
               sql = "select OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OSLP', 'SlpName', SlpCode, SlpName) SlpName from oslp where SlpCode = " & SlpCode
               set rc = conn.execute(sql) %><font size="1" face="Verdana"><%=rc("SlpName")%></font><input type="hidden" name="SlpCode" value="<%=SlpCode%>"><% End If %></td>
            </tr>
            <tr>
              <td width="33%" bgcolor="#7DB1FF"><b><font size="1" face="Verdana"><%=getcart_extraLngStr("DtxtProject")%></font></b></td>
              <td width="67%" colspan="2">
              <% If myApp.EnableDocPrjSel Then %>
		        <select size="1" name="Project" style="font-size:10px; width:100%; font-family:Verdana">
		        <option></option>
		        <%  
				set cmd = Server.CreateObject("ADODB.Command")
				cmd.ActiveConnection = connCommon
				cmd.CommandType = &H0004
				cmd.CommandText = "DBOLKGetProjects" & Session("ID")
				cmd.Parameters.Refresh()
				cmd("@LanID") = Session("LanID")
				set rc = cmd.execute()
		        do while not rc.eof %>
               <option <% If Project = rc("PrjCode") then response.write "selected"%> value="<%=rc("PrjCode")%>"><%=myHTMLEncode(rc("PrjName"))%></option>
               <%rc.movenext
               loop %>
               </select><%
               Else %><input type="hidden" name="Project" value="<%=PrjCode%>"><% End If %></td>
            </tr>
            <tr>
            	<td width="33%" bgcolor="#7DB1FF"><b><font size="1" face="Verdana"><%=getcart_extraLngStr("LtxtPymntCod")%></font></b></td>
            	<td width="67%" colspan="2"><% If myAut.HasAuthorization(88) Then %><select size="1" name="GroupNum" style="font-size:10px; width:100%; font-family:Verdana" onchange="checkPList(this.selectedIndex);">
				    <% 
				    sql = "select GroupNum, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCTG', 'PymntGroup', GroupNum, PymntGroup) PymntGroup, ListNum from octg"
				    set rc = Server.CreateObject("ADODB.RecordSet")
				    rc.open sql, conn, 3, 1
				    do while not rc.eof %>
			        <option value="<%=rc("GroupNum")%>" <% If CInt(rc("GroupNum")) = GroupNum then response.write "selected" %>><%=myHTMLEncode(rc("PymntGroup"))%></option>
			        <% rc.movenext
			        loop
			        rc.movefirst %>
			        </select>
			        <script language="javascript">
			        var ListNum = new Array(<%=rc.recordcount%>);
			        <% do while not rc.eof %>
			        ListNum[<%=rc.bookmark-1%>] = <%=rc("ListNum")%>;<%
			        rc.movenext
			        loop %>
			        </script><% Else
			        sql = "select OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCTG', 'PymntGroup', GroupNum, PymntGroup) PymntGroup, ListNum from octg where GroupNum = " & GroupNum
			        set rc = conn.execute(sql) %><font size="1" face="Verdana"><%=rc("PymntGroup")%></font>
			        <script language="javascript">
			        var ListNum = new Array(1);
			        ListNum[0] = <%=rc("ListNum")%>;
			        </script><input type="hidden" name="GroupNum" value="<%=GroupNum%>"><% End If %><input type="hidden" name="ChangePList" value="N"><input type="hidden" name="NewPriceList" value="<%=Session("plist")%>">
				</td>
            </tr>
            <% do while not rg.eof %>
            <tr>
              <td colspan="3" bgcolor="#7DB1FF"><b><font size="1" face="Verdana"><% Select Case CInt(rg("GroupID"))
              Case -1 %><%=getcart_extraLngStr("DtxtUDF")%><%
              Case Else
              	Response.Write rg("GroupName")
              End Select %></font></b></td>
            </tr>
			<% rcOpt.Filter = "GroupID = " & rg("GroupID")
			do while not rcOpt.eof 
            AliasID = rcOpt("InsertID")
            If Request.Form.Count = 0 Then
	            fldVal = rcOptVals(AliasID)
	            If rcOpt("TypeID") = "D" Then fldVal = FormatDate(fldVal, False)
	        Else
	        	fldVal = Request("U_" & rcOpt("AliasID"))
	        End If %>
            <tr>
              <td width="33%" bgcolor="#7DB1FF"><b>
                      <font size="1" face="Verdana">&nbsp;<%=rcOpt("Descr")%><% If rcOpt("NullField") = "Y" Then %><font color="red">*</font><% End If %></font></b></td>
              <td width="67%" colspan="2">
        		<% If rcOpt("DropDown") = "Y" or Not IsNull(rcOpt("RTable")) then 
        		If rcOpt("DropDown") = "Y" Then
	        		sql = "select FldValue, IsNull(AlterDescr, Descr) Descr " & _
									"from UFD1 T0 " & _
									"left outer join OLKUFD1AlterNames T1 on T1.TableID = T0.TableID and T1.FieldID = T0.FieldID and T1.IndexID = T0.IndexID and T1.LanID = " & Session("LanID") & " " & _
									"where T0.tableid = 'OINV' and T0.FieldId = " & rcOpt("FieldId")
				Else
					sql = "select Code FldValue, Name Descr from [@" & rcOpt("RTable") & "] order by 2"
				End If
				set rctn = conn.execute(sql) %>
				<font color="#4783C5">
				<select size="1" name="U_<%=rcOpt("AliasID")%>" class="input" style="font-size:10px; width:100%; font-family:Verdana">
				<option></option>
				<% do while not rctn.eof %>
				<option value="<%=rctn("FldValue")%>" <% If fldVal = rctn("FldValue") Then %>selected<% ElseIf rctn("FldValue") = rcOpt("Dflt") and IsNull(fldVal) Then %>selected<% End If %>><%=rctn("Descr")%></option>
				<% rctn.movenext
				loop
				rctn.close %></select></font><font size="1" color="#4783C5">
				<% Else %>
				<% If rcOpt("Query") = "Y" or rcOpt("TypeID") = "D" Then %><table width="100%" cellspacing="0" cellpadding="0"><tr><td width="16"><a href="#" <% If rcOpt("TypeID") = "D" Then %>onclick="javascript:getCal('<%=rcOpt("AliasID")%>', 'N')"<% End If %> <% If rcOpt("Query") = "Y" Then %>onclick="javascript:getVal('<%=rcOpt("AliasID")%>')"<% End If %>><img border="0" src="<% If rcOpt("Query") = "Y" Then %>../images/<%=Session("rtl")%>flechaselec2.gif<% Else %>images/cal.gif<% End If %>"></a></td><td><% End If %>
				<input <% If rcOpt("Query") = "Y" or rcOpt("TypeID") = "D" Then %>readonly<% End If %> type="<% If rcOpt("TypeID") = "N" or rcOpt("TypeID") = "B" Then %>number<% Else %>text<% End If %>" <% If rcOpt("TypeID") = "A" Then %>maxlength="<%=rcOpt("SizeID")%>"<% End If %> name="U_<%=rcOpt("AliasID")%>" size="<% If rcOpt("TypeID") = "A" Then %>43<% Else %>12<% End If %>"  class="input" value="<% If fldVal <> "" Then %><%=fldVal%><% Else %><%=rcOpt("Dflt")%><% End If %>" onclick="this.selectionStart=0;this.selectionEnd=this.value.length;" style="width: 100%; font-family: Verdana; font-size: 10px">
				<% If rcOpt("Query") = "Y" or rcOpt("TypeID") = "D" Then %></td><td width="16"><a href="#" onclick="javascript:document.frmCart.U_<%=rcOpt("AliasID")%>.value = ''"><img border="0" src="../images/remove.gif" width="16" height="16"></a></td></tr></table><% End If %><% End If %></td>
            </tr>
			<% 
			rcOpt.movenext
			loop 
			rg.movenext
			loop
			rcOpt.Filter = "" %>
            </table>
          </td>
        </tr>
        <tr>
          <td width="100%">
          <table border="0" cellpadding="0" cellspacing="1"  bordercolor="#111111" width="100%" height="60">
            <tr>
              <td width="100%" height="12" bgcolor="#7DB1FF"><b>
              <font size="1" face="Verdana"><%=getcart_extraLngStr("LtxtComments")%></font></b></td>
            </tr>
            <tr>
              <td width="100%" height="36" bgcolor="#95BFFF">
              <p align="center"><textarea rows="2" name="Comments" cols="23"><%=myHTMLEncode(Comments)%></textarea></td>
            </tr>
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
                <a href="#" onclick="javascript:ValidateForm();"><img src="images/save_icon.gif" border="0" alt=""></a>
                </td>
                <td>
                <p align="center"><a href="operaciones.asp?cmd=cart"><img border="0" src="images/x_icon.gif"></a></td>
              </tr>
            </table>
            </center>
          </div>
          </td>
        </tr>
      </table>
      	<input type="hidden" name="cmd" value="cartExtra">
      	<input type="hidden" name="editVar" value="">
      	<input type="hidden" name="returnCmd" value="cartopt">
      </form>
      </td>
    </tr>
    </table>
  </center>
</div>
<% set rc = nothing 
set rctn = nothing
set rcOpt = nothing
set rcOptVals = nothing
If Request("compFld") = "Y" Then %>
<script language="javascript">alert('<%=getcart_extraLngStr("DtxtCompFld")%>');</script>
<% End If %>
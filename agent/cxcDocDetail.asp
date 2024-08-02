<!--#include file="lang/cxcDocDetail.asp" -->
<!--#include file="myHTMLEncode.asp"-->
<!--#include file="authorizationClass.asp"-->
<html <% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>>
<%
Dim varx
varx = "0"
%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" >
<title><%=getcxcDocDetailLngStr("LtxtDocDet")%></title>
<link type="text/css" href="design/0/jquery-ui-1.7.2.custom.css" rel="stylesheet" >	
<script type="text/javascript" src="jQuery/js/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="jQuery/js/jquery-ui-1.7.2.custom.min.js"></script>
<script type="text/javascript" src="general.js"></script>
<!--#include file="getNumeric.asp"-->
<%
Dim myAut
set myAut = New clsAuthorization

If Request("Excell") = "Y" Then response.ContentType="application/vnd.ms-excel"

DocType = CInt(Request("DocType"))

Dim DocName
set rs = Server.CreateObject("ADODB.recordset")
sql = "select top 1 SelDes, VatPrcnt from oadm cross join OLKCommon order by CurrPeriod desc"
set rs = conn.execute(sql)
VatPrcnt = rs("VatPrcnt")
If userType = "V" Then SelDes = 0 Else SelDes = rs("SelDes")
ShowClientRef = myApp.ShowClientRef or userType = "V"

If userType = "V" Then
	EnableDiscount = myAut.HasAuthorization(68)
	PrintPriceBefDiscount = myApp.PrintPriceBefDiscount
	PrintLineDiscount = myApp.PrintLineDiscount
End If %>
<!--#include file="loadAlterNames.asp" -->
<!--#include file="clearItem.asp" -->
<%    
          
If Request("PDF") = "Y" and userType = "V" Then
	myAut.LoadAuthorization Request("vendid"), ""
	sql = "select Access from OLKAgentsAccess where SlpCode = " & Request("vendid")
	set rs = conn.execute(sql)
	Session("useraccess") = rs("Access")
End If
           
        set rd = Server.CreateObject("ADODB.RecordSet")
        set rdocnum = Server.CreateObject("ADODB.recordset")
        If Request("isEntry") <> "N" Then isEntry = "Y" Else isEntry = "N"
           
        set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKCXCDocInfo" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@DocType") = CLng(Request("DocType"))
		cmd("@DocEntry") = CLng(Request("DocEntry"))
		cmd("@LanID") = CInt(Session("LanID"))
		cmd("@isEntry") = isEntry
        set rs = cmd.execute()
        
		If Not rs.Eof Then
			chkType = CInt(rs("ObjType"))
			If chkType = 13 Then If rs("ReserveInvoice") = "Y" Then chkType = -13
			hasAut = userType = "C" or myAut.GetObjectProperty(chkType, "V")
			hasClientAccess = userType = "C" or userType = "V" and (myAut.HasAuthorization(60) or not myAut.HasAuthorization(60) and CInt(rs("SlpCode")) = CInt(Session("vendid"))) 
			hasDocAut = userType = "C" or userType = "V" and (myAut.HasAuthorization(97) or not myAut.HasAuthorization(97) and CInt(rs("SlpCode")) = CInt(Session("vendid")))
		End If
		If hasAut and hasClientAccess and hasDocAut Then

           DocEntry = rs("DocEntry")
           If Not (userType = "V" or userType = "C" and (LCase(saveHTMLDecode(Session("username"), False)) = LCase(rs("CardCode")) or Not IsNull(rs("FatherCard")) and LCase(saveHTMLDecode(Session("username"), False)) = LCase(rs("FatherCard")))) Then 
           		Response.Redirect "noaccess.asp?ErrCode=0"
           End If

			If Request("DocType") <> "-2" Then
					sql= "select T1, T2 from olkDocConf where ObjectCode = " & Request("DocType")
					set rd = conn.execute(sql)
					oTable = rd("T1")
					oTable1 = rd("T2")
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
				  "where T0.TableId = 'OINV' and AType in ('" & userType & "', 'T') and OP in ('T','O') and Active = 'Y' " & _
				  "order by T3.[Order], IsNull(T1.Pos, 'D'), IsNull(T1.[Order], 32727) "
			rcOpt.open sql, conn, 3, 1

			set rctn = Server.CreateObject("ADODB.RecordSet")
			sql = "select AliasID, NullField, IsNull(AlterDescr, Descr) Descr " & _
				  "from cufd T0 " & _
				  "left outer join OLKCUFD T1 on T1.TableID = T0.TableID and T1.FieldID = T0.FieldID " & _
				  "left outer join OLKCUFDAlterNames T2 on T2.TableID = T0.TableID and T2.FieldID = T0.FieldID and T2.LanID = " & Session("LanID") & " " & _
				  "where T0.TableId = 'OINV' and AType in ('" & userType & "', 'T') and OP in ('T','O') and Active = 'Y' and NullField = 'Y' " & _
				  "Order By Pos Desc"
			rctn.open sql, conn, 3, 1
			
			If rctn.recordcount > 0 Then chkOpt = True
			
			If rcOpt.RecordCount > 0 Then
				set rcOptVals = Server.CreateObject("ADODB.RecordSet")
				sql = "select " & sqlAddStr & "++AliasID As InsertID, T0.FieldID, " & _
					  "Case When Exists(select 'A' from UFD1 where TableId = T0.TableId and FieldId = T0.FieldId) " & _
				  	  "Then 'Y' Else 'N' End As DropDown, RTable, TypeID " & _
				  	  "from cufd T0 " & _
				  	  "left outer join OLKCUFD T1 on T1.TableID = T0.TableID and T1.FieldID = T0.FieldID " & _
					  "where T0.TableId = 'OINV' and AType in ('" & userType & "','T') and OP in ('T','O') and Active = 'Y'"
				rcOptVals.open sql, conn, 3, 1
				sql = "select "
				do while not rcOptVals.eof
					If rcOptVals.bookmark <> 1 Then sql = sql & ", "
					If rcOptVals("TypeID") <> "N" and rcOptVals("TypeID") <> "B" Then strCollate = "collate database_default" Else strCollate = ""
					If rcOptVals("DropDown") = "Y" and (IsNull(rcOptVals("RTable")) or rcOptVals("RTable") = "") Then
						sql = sql & "(select IsNull((select AlterDescr from OLKUFD1AlterNames where TableID = UFD1.TableID and FieldID = UFD1.FieldID and IndexID = UFD1.IndexID and LanID = " & Session("LanID") & "), Descr) from UFD1 where TableID = 'OINV' and FieldID = " & rcOptVals("FieldID") & " and FldValue = T0." & rcOptVals("InsertID") & " " & strCollate & ") '" & rcOptVals("InsertID") & "'"
					ElseIf (Not IsNull(rcOptVals("RTable")) and rcOptVals("RTable") <> "") Then
						sql = sql & "(select Name from [@" & rcOptVals("RTable") & "] where Code = T0." & rcOptVals("InsertID") & " " & strCollate & ") '" & rcOptVals("InsertID") & "'"
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
			
			addColSpan = 0
			If ShowClientRef Then addColSpan = addColSpan + 1
			If myApp.GetShowSalUn Then addColSpan = addColSpan + 1
			If EnableDiscount Then
				If PrintLineDiscount Then addColSpan = addColSpan + 1
				If PrintPriceBefDiscount Then addColSpan = addColSpan + 1
			End If 
			
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
				
				set rx = Server.CreateObject("ADODB.RecordSet")
				DocEntryFld = "LogNum"
				QueryFld = "RowQuery"
				QueryTable = "DOC"
				If Request("doctype") <> "-2" Then
					DocEntryFld = "DocEntry"
					QueryFld = "SystemQuery"
					sql = "select Right(T1, 3) from OLKDocConf where ObjectCode = " & Request("doctype")
					set rx = conn.execute(sql)
					If Not rx.Eof Then QueryTable = rx(0) Else QueryTable = "Disable"
					rx.close
				End If
				
				bottomRowSpan = 8
				
				If QueryTable <> "Disable" Then
				
				sql = "select IsNull(T1.AlterRowName, T0.RowName) RowName, T0." & QueryFld & " [Query] " & _
				"from OLKCMREP T0 " & _
				"left outer join OLKCMREPAlterNames T1 on T1.RowType = T0.RowType and T1.LineIndex = T0.LineIndex and T1.LanID = " & Session("LanID") & " " & _
				"where T0.RowActive = 'Y' and T0.Print" & userType & " = 'Y' and " & QueryFld & " is not null " & _
				"order by T0.RowOrder asc"
				
				rx.open sql, conn, 3, 1
				bottomRowSpan = bottomRowSpan + 1 + rx.recordcount 
				
				End If

If Request("Excell") <> "Y" and Request("PDF") <> "Y" Then %>
<script type="text/javascript">
<!--
function saveDoc(cmd)
{
	switch (cmd)
	{
		case 'Print':
			document.getElementById('tblSave').style.display = 'none';
			document.getElementById('crdLink').style.display = 'none';
			showItmLink(false);
			window.print();
			showItmLink(true);
			document.getElementById('tblSave').style.display = '';
			document.getElementById('crdLink').style.display = '';
			break;
		case 'PDF':
			document.frmExcell.action = 'cxcDocDetailPDF.asp';
			document.frmExcell.submit();
			break;
		case 'Excell':
			document.frmExcell.action = 'cxcDocDetailOpen.asp';
			document.frmExcell.submit();
			break;
	}
}
function showItmLink(show)
{
	if (document.getElementById('itmLink') != null)
	{
		myLinks = document.getElementById('itmLink');
		if (myLinks.length)
		{
			for (var i = 0;i<myLinks.length;i++)
			{
				myLinks[i].style.display = show ? '' : 'none';
			}
		}
		else
		{
			myLinks.style.display = show ? '' : 'none';
		}
	}
}

//-->
</script>
<link rel="stylesheet" type="text/css" href="design/<%=SelDes%>/style/stylenuevo.css">
<% End If %>
</head>
<body>
<% If Request("Excell") <> "Y" and Request("PDF") <> "Y" Then %>
<table border="0" cellpadding="0" cellspacing="0" width="100%" id="tblSave">
	<tr>
		<td align="right">
		<a href="#" onclick="javascript:saveDoc('Print');"><img alt="<%=getcxcDocDetailLngStr("DtxtPrint")%>" border="0" src="images/print_OLK.gif"></a>&nbsp;
		<% If userType = "C" or userType = "V" and myAut.HasAuthorization(65) Then %><a href="#" onclick="javascript:saveDoc('PDF');"><img alt="<%=getcxcDocDetailLngStr("DtxtExpPDF")%>" border="0" src="images/pdf_OLK.gif"></a>&nbsp;<% End If %>
		<% If userType = "C" or userType = "V" and myAut.HasAuthorization(64) Then %><a href="#" onclick="javascript:saveDoc('Excell');"><img alt="<%=getcxcDocDetailLngStr("LtxtExpToExcell")%>" border="0" src="images/excell.gif"></a><% End If %>
		</td>
	</tr>
</table>
<% End If %>
<table border="0" cellpadding="0" width="100%">
	<% 
	set rsdf = Server.CreateObject("ADODB.RecordSet")
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKGetUDFSystemCols" & Session("ID")
	cmd.Parameters.Refresh
	cmd("@LanID") = Session("LanID")
	cmd("@UserType") = userType
	cmd("@TableID") = "OINV"
	cmd("@OP") = "O"
	rsdf.open cmd, , 3, 1
	If rs("PrintCmpPaper") = "Y" Then %>
	<tr class="FirmTlt">
		<td height="80">&nbsp;</td>
	</tr>
	<% Else %>
	<%=strContent%>
	<!--#include file="myTitle.asp"-->
	<% End If %>
	<% If rs("Draft") = "Y" or rs("Confirmed") = "N" Then %>
	<tr>
		<td id="tdMyTtl" class="TablasTituloDraft">
		<% If rs("Draft") = "Y" Then %><%=getcxcDocDetailLngStr("LttlDraftNote")%><% ElseIf rs("Confirmed") = "N" Then %><%=getcxcDocDetailLngStr("LttlConfirmNote")%><% End If %></td>
	</tr>
	<% End If %>
	<tr class="CanastaTitle">
		<td>
		<%=rs("ObjDesc")%><% If Request("DocType") = "-2" and rs("ObjDesc") <> rs("ObjectName") Then %>&nbsp;(<%=rs("ObjectName")%>)<% End If %>&nbsp;#<%=rs("DocNum")%></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" cellspacing="1">
			<tr>
				<td width="50%" valign="top">
				<table border="0" cellpadding="0" width="100%">
					<tr>
						<td class="CanastaTblResaltada" width="95"><%=getcxcDocDetailLngStr("DtxtCode")%></td>
						<td class="CanastaTbl"><%=RS("CardCode")%>&nbsp;</td>
					</tr>
					<tr>
						<td class="CanastaTblResaltada" width="95">
						<table cellpadding="0" cellspacing="0" border="0" width="100%">
							<tr class="CanastaTblResaltada">
								<td><%=getcxcDocDetailLngStr("LtxtTo")%></td>
								<% If Request("Excell") <> "Y" and Request("PDF") <> "Y" Then %><td align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
								<a href="javascript:goDetail('<%=rs("CardCode")%>')"><img id="crdLink" src="design/<%=SelDes%>/images/<%=Session("rtl")%>flecha_selec.gif" border="0"></a></td><% End If %>
							</tr>
						</table></td>
						<td class="CanastaTbl"><%=RS("cardname")%></td>
					</tr>
					<tr>
						<td class="CanastaTblResaltada" width="95" valign="top">
						<nobr><%=getcxcDocDetailLngStr("LtxtShipAdd")%></nobr></td>
						<td class="CanastaTbl">
						<%=RS("ShipToCode")%>:<br>
						<%=RS("ShipAddress")%>&nbsp;</td>
					</tr>
					<tr>
						<td class="CanastaTblResaltada" width="95" valign="top">
						<nobr><%=getcxcDocDetailLngStr("LtxtPayAdd")%></nobr></td>
						<td class="CanastaTbl">
						<%=RS("PayToCode")%>:<br>
						<%=RS("PayAddress")%>&nbsp;</td>
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
					<% 
					rsdf.Filter = "FieldID = -1"
					If Not rsdf.eof Then
					%>
					<tr>
						<td class="CanastaTblResaltada" width="95">
						<%=getcxcDocDetailLngStr("DtxtContact")%></td>
						<td class="CanastaTbl">
						<%=RS("Name")%>&nbsp;&nbsp;</td>
					</tr>
					<% End If %>

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
				<table border="0" cellpadding="0" width="100%">
					<tr>
						<td class="CanastaTblResaltada" width="95"><%=getcxcDocDetailLngStr("DtxtDate")%></td>
						<td class="CanastaTbl">
						<%=FormatDate(rs("DocDate"), True)%>&nbsp;</td>
					</tr>
					<% 
					rsdf.Filter = "FieldID = -3"
					If Not rsdf.eof Then
					%>
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
						<%=FormatDate(rs("DocDueDate"), True)%>&nbsp;</td>
					</tr>
					<% End If 
					rsdf.Filter = "FieldID = -2"
					If Not rsdf.eof Then
					%>
					<tr>
						<td class="CanastaTblResaltada" width="95"><% If 1 = 2 Then %>txtRef2<% Else %><%=Server.HTMLEncode(txtRef2)%><% End If %></td>
						<td class="CanastaTbl"><%=RS("NumAtCard")%>&nbsp;</td>
					</tr>
					<% End If %>
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
                      <% ElseIf rcOpt("TypeID") = "D" Then %>
                      <%=FormatDate(rcOptVals(AliasID), True)%>
                      <% Else %>
                      <% If Not IsNull(rcOptVals(AliasID)) Then %><%=rcOptVals(AliasID)%><% End If %>
                      <% End If %>
                      <% If rcOpt("TypeID") = "M" and rcOpt("EditType") = "B" Then %></a><% End If %>
                      </td>
                    </tr>
                    <% 
                    rcOpt.movenext
                    loop  
                    set rcOpt = Nothing %>
					</table>
				</td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr class="FirmTlt3">
				<td>#</td>
				<% If ShowClientRef Then %><td align="center"><%=getcxcDocDetailLngStr("DtxtCode")%></td><% End If %>
				<td align="center"><%=getcxcDocDetailLngStr("LtxtProd")%></td>
				<td align="center"><%=getcxcDocDetailLngStr("DtxtQty")%></td>
				<% If myApp.GetShowSalUn Then %><td align="center"><%=getcxcDocDetailLngStr("DtxtUnit")%></td><% End If %>
				<% If EnableDiscount Then %>
				<% If PrintPriceBefDiscount Then %>
				<td align="center"><%=getcxcDocDetailLngStr("LtxtUnitPrice")%></td>
				<% End If %>
				<% If PrintLineDiscount Then %>
				<td align="center"><%=getcxcDocDetailLngStr("DtxtDiscount")%></td>
				<% End If %>
				<% End If %>
				<td align="center"><% If Not EnableDiscount or EnableDiscount and not (PrintPriceBefDiscount or PrintLineDiscount) Then %><%=getcxcDocDetailLngStr("DtxtPrice")%><% Else %><%=getcxcDocDetailLngStr("LtxtPriceAfterDisc")%><% End If %></td>
				<td align="center"><%=getcxcDocDetailLngStr("DtxtTotal")%></td>
			</tr>
			  <% 
			  set cmd = Server.CreateObject("ADODB.Command")
			  cmd.ActiveConnection = connCommon
			  cmd.CommandType = &H0004
			  cmd.CommandText = "DBOLKGetDocLinesDetails" & Session("ID")
			  cmd.Parameters.Refresh()
			  cmd("@LanID") = Session("LanID")
			  cmd("@DocType") = DocType
			  cmd("@SlpCode") = Session("vendid")
			  If DocType = -2 Then
			  	cmd("@LogNum") = DocEntry
			  Else
			  	cmd("@DocEntry") = DocEntry
			  End If
			  set rd = Server.CreateObject("ADODB.RecordSet")
			  rd.open cmd, , 3, 1
			  lNum = 0
			  do while not rd.eof 
			  lNum = lNum + 1
			  If Request("high") <> "" Then
				  If CStr(rd("LineNum")) = CStr(Request("high")) or CStr(rd("LineNum")) = CStr(Request("LineNum")) Then
				  	HighLight = True
				  Else
				  	HighLight = False
				  End If
			  Else
			  	HighLight = False
			  End If
			  
			  TreeType = rd("TreeType")
			  TreePricOn = rd("TreePricOn") = "Y"
			  ShowPriceAndTotal = (TreeType <> "S" and TreeType <> "I" and TreeType <> "C") or (((TreeType = "S" and (IsNull(rd("ShowFatherPrice"))) or rd("ShowFatherPrice") = "Y") or TreeType = "I" or TreeType = "C") and rd("Currency") <> "" and CDbl(rd("LineTotal")) <> 0)
			  			  %>
			<tr class="CanastaTbl" style="<% If HighLight then %>color: #BB0000;<% End If %><% Select Case rd("TreeType")
				Case "S" %>font-weight: bold;<%
				Case "I" %>font-style: italic;<%
			End Select %>">
				<td><%=lNum%></td>
				<% If ShowClientRef Then %><td><% If Request("Excell") <> "Y" and Request("PDF") <> "Y" Then %><a href="javascript:goItem('<%=rd("ItemCode")%>', '<%=rd("WhsCode")%>')"><img id="itmLink" src="design/<%=SelDes%>/images/<%=Session("rtl")%>flecha_selec.gif" border="0"></a><% End If %><%=RD("ItemCode")%>&nbsp;</td><% End If %>
				<td><% If Not ShowClientRef Then %><% If Request("Excell") <> "Y" and Request("PDF") <> "Y" Then %><a href="javascript:goItem('<%=rd("ItemCode")%>', '<%=rd("WhsCode")%>')"><img id="itmLink" src="design/<%=SelDes%>/images/<%=Session("rtl")%>flecha_selec.gif" border="0"></a><% End If %><% End If %><%=RD("ItemName")%>&nbsp;</td>
				<td>
				<p align="right">&nbsp;<%=FormatNumber(Cdbl(RD("Quantity")),myApp.QtyDec)%></td>
				<% If myApp.GetShowSalUn Then %><td align="center"><%=RD("SalUnitMsr")%>&nbsp;</td><% End If %>
				<% If EnableDiscount Then %>
				<% If PrintPriceBefDiscount Then %>
				<td>
				<p align="right"><% If ShowPriceAndTotal Then %><nobr><%=rd("Currency")%>&nbsp;<%=FormatNumber(CDbl(RD("UnitPrice")),myApp.PriceDec)%></nobr><% Else %>&nbsp;<% End If %></td>
				<% End If %>
				<% If PrintLineDiscount Then %>
				<td dir="ltr">
				<p align="right"><nobr>%&nbsp;<%=FormatNumber(CDbl(RD("DiscPrcnt")),myApp.PercentDec)%></nobr></td>
				<% End If %>
				<% End If %>
				<td>
				<p align="right"><% If ShowPriceAndTotal Then %><nobr><%=rd("Currency")%>&nbsp;<%=FormatNumber(RD("Price"),myApp.PriceDec)%></nobr><% Else %>&nbsp;<% End If %></td>
				<td>
				<p align="right"><% If ShowPriceAndTotal Then %><nobr><%=Rs("DocCur")%>&nbsp;<%=FormatNumber(RD("LineTotal"),myApp.SumDec)%></nobr><% Else %>&nbsp;<% End If %></td>
			</tr>
		    <% Rd.MoveNext
			loop %>
			<tr>
				<td colspan="<%=3+addColSpan%>" rowspan="<%=bottomRowSpan%>" valign="top">
				<div align="left">
				<% If rs("PayLogNum") <> "" or Request("payment") = "Y" Then 
				If rs("PayLogNum") <> "" Then PayLogNum = rs("PayLogNum") Else PayLogNum = Session("ConfPayRetVal")
				
				sql = "declare @LogNum int set @LogNum = " & PayLogNum & " " & _
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
				<table border="0" cellpadding="0" width="80%" cellspacing="1">
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

		<table border="0" cellpadding="0" width="100%" cellspacing="1">
			<tr>
				<td width="261" valign="top">
				<table border="0" cellpadding="0" width="100%">
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
				<p align="right"><nobr><%=rs("DocCur")%>&nbsp;<%=FormatNumber(Rs("SubTotal"),myApp.SumDec)%></nobr></td>
			</tr>
			<% If rs("ObjType") <> 203 and rs("ObjType") <> 204 Then %>
			<tr>
				<td class="CanastaTblResaltada">
				<table cellpadding="0" cellspacing="0" border="0" width="100%">
					<tr class="CanastaTblResaltada">
						<td><nobr><%=getcxcDocDetailLngStr("DtxtDiscount")%></nobr></td>
						<td style="font-weight: normal;" align="right"><nobr>%&nbsp;<%=FormatNumber(CDbl(rs("DiscPrcnt")),myApp.PercentDec)%></nobr></td>
					</tr>
				</table>
				</td>
				<td class="CanastaTbl">
				<p align="right" dir="ltr"><nobr><%=rs("DocCur")%>&nbsp;<%=FormatNumber((CDbl(Rs("SubTotal"))*(CDbl(rs("DiscPrcnt"))/100)),myApp.SumDec)%></nobr></td>
			</tr>
			<% Else %>
			<tr>
				<td class="CanastaTblResaltada">
				<table cellpadding="0" cellspacing="0" border="0" width="100%">
					<tr class="CanastaTblResaltada">
						<td><nobr><%=getcxcDocDetailLngStr("DtxtDPM")%></nobr></td>
						<td style="font-weight: normal;" align="right"><nobr>%&nbsp;<%=FormatNumber(CDbl(rs("DpmPrcnt")),myApp.PercentDec)%></nobr></td>
					</tr>
				</table>
				</td>
				<td class="CanastaTbl">
				<p align="right" dir="ltr"><nobr><%=rs("DocCur")%>&nbsp;<%=FormatNumber(CDbl(rs("DpmAmnt")),myApp.SumDec)%></nobr></td>
			</tr>
			<% End If %>
          <%
            If myApp.ExpItems Then
				cmd.CommandText = "DBOLKGetDocExpnsDetails" & Session("ID")
				cmd.Parameters.Refresh()
				cmd("@DocType") = DocType
				If DocType = -2 Then cmd("@LogNum") = DocEntry Else cmd("@DocEntry") = DocEntry
				rd.close
				rd.open cmd, , 3, 1
			  do while not rd.eof
			%>
			<tr>
				<td class="CanastaTblResaltada"><%=rd("ItemName")%>&nbsp;</td>
				<td class="CanastaTbl">
				<p align="right"><nobr><%=Rd("Currency")%>&nbsp;<%=FormatNumber(Rd("LineTotal"),myApp.SumDec)%></nobr></td>
			</tr>
			<% rd.movenext
			loop 
			End If %>
			<tr>
				<td class="CanastaTblResaltada"><% If 1 = 2 Then %>txtTax<% Else %><%=txtTax%><% End If %><% If myApp.LawsSet = "IL" Then %><% If Session("myLng") = "he" Then %><span style="font-size: xx-small; "><% End If %>&nbsp;%<%=FormatNumber(VatPrcnt, myApp.PercentDec)%><% If Session("myLng") = "he" Then %></span><% End If %><% End If %></td>
				<td class="CanastaTbl">
				<p align="right"><nobr><%=RS("DocCur")%>&nbsp;<%=FormatNumber(Rs("ITBM"),myApp.SumDec)%></nobr></td>
			</tr>
			<tr>
				<td class="CanastaTblResaltada"><%=getcxcDocDetailLngStr("DtxtTotal")%></td>
				<td class="CanastaTbl">
				<p align="right"><nobr><%=RS("DocCur")%>&nbsp;<%=FormatNumber(Rs("DocTotal"),myApp.SumDec)%></nobr></td>
			</tr>
			<% If Not IsNull(rs("PayLogNum")) or Request("DocType") = 13 or Request("DocType") = 18 Then %>
			<tr>
				<td class="CanastaTblResaltada"><%=getcxcDocDetailLngStr("DtxtPaid")%></td>
				<td class="CanastaTbl"><p align="right">
				<nobr><% If Not IsNull(rs("PayLogNum")) Then %>
				<%=PayDocCur%>&nbsp;<%=FormatNumber(SumApplied,myApp.SumDec)%>
				<% Else %>
				<%=rs("DocCur")%>&nbsp;<%=FormatNumber(CDbl(rs("PaidToDate")), myApp.SumDec)%>
				<% End If %></nobr></td>
			</tr>
			<tr>
				<td class="CanastaTblResaltada"><%=getcxcDocDetailLngStr("DtxtBalance")%></td>
				<td class="CanastaTbl"><p align="right">
				 <nobr><% If Not IsNull(rs("PayLogNum")) Then
				 OpenSum = CDbl(Rs("DocTotal"))-CDbl(SumApplied)
                  If OpenSum < 0 Then OpenSum = 0 %>
                  <%=RS("DocCur")%>&nbsp;<%=FormatNumber(OpenSum,myApp.SumDec)%>
                  <% Else %>
                  <%=rs("DocCur")%>&nbsp;<%=FormatNumber(CDbl(rs("OpenBalance")), myApp.SumDec)%>
                  <% End If %></nobr></td>
			</tr>
			<% End If 
				
				If QueryTable <> "Disable" Then
				If Not rx.Eof Then %>
				<tr>
					<td>&nbsp;</td>
					<td>&nbsp;</td>
				</tr>
				<% sql = "declare @LanID int set @LanID = " & Session("LanID") & " declare @" & DocEntryFld & " int set @" & DocEntryFld & " = " 
				If QueryTable = "DOC" or QueryTable <> "DOC" and isEntry = "Y" Then
					sql = sql & Request("DocEntry") & " "
				Else
					sql = sql & "(select DocEntry from O" & QueryTable & " where DocNum = " & Request("DocEntry") & ") "
				End If

				sql = sql & "declare @CardCode nvarchar(15) set @CardCode = N'" & saveHTMLDecode(rs("CardCode"), False) & "' " & _
				"select "
				do while not rx.eof
					If rx.bookmark > 1 Then sql = sql & ", "
					sql = sql & "(" & rx("Query") & ") As '" & rx("RowName") & "'"
				rx.movenext
				loop
				If QueryTable <> "DOC" Then sql = Replace(sql, "{Table}", QueryTable)
				sql = QueryFunctions(sql)
				set rx = conn.execute(sql)
				For each fld in rx.Fields
				%>
				<tr>
					<td class="CanastaTblResaltada"><%=fld.Name%>&nbsp;</td>
					<td class="CanastaTbl" align="right" dir="ltr">&nbsp;<%=fld%></td>
				</tr>
				<% Next %>
				<% 
				End If
				End If %>
			</table>
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
</table>
<% If Not IsNull(RS("ObjectNote")) Then %>
<%=RS("ObjectNote")%>
<% End If %>

<form target="_blank" method="post" name="frmViewDetail" action="addCard/crdConfDetailOpen.asp">
<input type="hidden" name="DocEntry" value="">
<input type="hidden" name="DocType" value="2">
<input type="hidden" name="CardCode" value="">
<input type="hidden" name="pop" value="Y">
</form>
<% If userType = "C" Then %>
<form name="frmGoItem" action="item.asp" method="post">
<input type="hidden" name="Item" value="">
<input type="hidden" name="WhsCode" value="">
<input type="hidden" name="cmd" value="d">
<input type="hidden" name="pop" value="Y">
<input type="hidden" name="AddPath" value="">
</form>
<% Else %>
<!--#include file="itemDetails.inc"-->
<% End If %>
<script language="javascript">
function goDetail(CardCode) {
	document.frmViewDetail.CardCode.value = CardCode;
	document.frmViewDetail.submit();
}
var rtl = '<%=Session("rtl")%>';
function goItem(ItemCode, WhsCode)
{
	<% Select Case userType
	Case "C" %>
	document.frmGoItem.Item.value = ItemCode;
	document.frmGoItem.WhsCode.value = WhsCode;
	document.frmGoItem.submit();
	<% Case "V" %>
	itemLoadWhs = WhsCode;
	openItemDetails(ItemCode);
	<% End Select %>
}
</script>

<% If Request("Excell") <> "Y" and Request("PDF") <> "Y" Then %>
<form name="frmExcell" method="post">
<% For each itm in Request.Form
If itm <> "CardCode" Then %><input type="hidden" name="<%=itm%>" value="<%=Request.Form(itm)%>"><% End If
Next %>
<input type="hidden" name="Excell" value="Y">
<% If userType = "C" Then %><input type="hidden" name="CardCode" value="<%=rs("CardCode")%>"><% End If %>
</form>
<% End If %>


<% Else %>
	<script type="text/javascript">
	<% If Not rs.Eof Then %>
	<% If Not hasAut Then %>
	alert('<%=getcxcDocDetailLngStr("DtxtNoAccessObj")%>'.replace('{0}', '<%=rs("ObjectName")%>'));
	<% ElseIf Not hasClientAccess Then %>
	alert('<%=getcxcDocDetailLngStr("DtxtNoClientAccess")%>'.replace('{0}', '<%=rs("CardCode")%>'));
	<% ElseIf Not hasDocAut Then %>
	alert('<%=getcxcDocDetailLngStr("DtxtNoDocAccess")%>');
	<% End If %>
	<% Else %>
	alert('<%=getcxcDocDetailLngStr("DtxtNoData")%>');
	<% End if %>
	window.close();
	</script>
<% End If %>
<!--#include file="linkForm.asp"-->
</body>
<% set rs = nothing
set rd = nothing
set rx = Nothing
conn.close %>
</html>
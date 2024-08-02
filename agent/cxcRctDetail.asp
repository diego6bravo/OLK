<html <% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>>
<% 
If Request("Excell") = "Y" Then response.ContentType="application/vnd.ms-excel"
 %>
<!--#include file="loadAlterNames.asp" -->  
<!--#include file="clearItem.asp" -->
<!--#include file="lang/cxcRctDetail.asp" -->     
<!--#include file="myHTMLEncode.asp"-->   
<head>
<title><%=getcxcRctDetailLngStr("LtxtPymntDet")%></title>
</head>
<body>
<%

		docType = CInt(Request("DocType"))
		If docType = 140 Then
			sql = "select ObjType from OPDF where DocEntry = " & Request("DocEntry")
			set rs = conn.execute(sql)
			docType = CInt(rs("ObjType"))
		End If

		Select Case docType
			Case 24
				hasAut = userType = "C" or myAut.GetObjectProperty(24, "V")
			Case 46
				hasAut = userType = "C" or myAut.GetObjectProperty(46, "V")
		End Select
		
		If hasAut Then

         set rc = Server.CreateObject("ADODB.recordset")
           	CmpName = mySession.GetCompanyName
			If userType = "V" Then 
				SelDes = 0 
			Else 
				sql = "select SelDes from OLKCommon"
				set rc = conn.execute(sql)
				SelDes = rc("SelDes")
			End If
           
If Request("PDF") = "Y" and userType = "V" Then
	myAut.LoadAuthorization Request("vendid"), ""
	sql = "select Access from OLKAgentsAccess where SlpCode = " & Request("vendid")
	set rs = conn.execute(sql)
	Session("useraccess") = rs("Access")
End If


set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKCXCRctInfo" & Session("ID")
cmd.Parameters.Refresh()
cmd("@DocType") = Request("DocType")
cmd("@DocEntry") = Request("DocEntry")
cmd("@MainCur") = myApp.MainCur
set rc = cmd.execute()
           
           If Not rc.Eof Then
           
           If Not (userType = "V" or userType = "C" and (LCase(saveHTMLDecode(Session("username"), False)) = LCase(rc("CardCode")) or Not IsNull(rc("FatherCard")) and LCase(saveHTMLDecode(Session("username"), False)) = LCase(rc("FatherCard")))) Then 
           		Response.Redirect "noaccess.asp?ErrCode=0"
           End If

			sqlInsertID = ""
			Select Case CInt(Request("DocType"))
				Case 24
					sqlInsertID = "'U_'"
					TableID = "ORCT"
					sqlDocNum = "DocEntry"
				Case 46
					sqlInsertID = "'U_'"
					TableID = "OVPM"
					sqlDocNum = "DocEntry"
				Case 140
					sqlInsertID = "'U_'"
					TableID = "OPDF"
					sqlDocNum = "DocEntry"
				Case Else
					sqlInsertID = "(select SDKID collate database_default from r3_obscommon..tcif where companydb = N'" & Session("OLKDb") & "')"
					TableID = "R3_ObsCommon..TPMT"
					sqlDocNum = "LogNum"
			End Select
			set rsUFD = Server.CreateObject("ADODB.RecordSet")
			sql = "select IsNull(AlterDescr, Descr) Descr, " & sqlInsertID & "++AliasID As InsertID, Pos, TypeID, EditType, " & _
				  "Pos, AliasID, Case When Exists(select 'A' from UFD1 where TableId = T0.TableId and FieldId = T0.FieldId) " & _
				  "Then 'Y' Else 'N' End As DropDown, T0.FieldID " & _
				  "from cufd T0 " & _
				  "left outer join OLKCUFD T1 on T1.TableID = T0.TableID and T1.FieldID = T0.FieldID " & _
				  "left outer join OLKCUFDAlterNames T2 on T2.TableID = T0.TableID and T2.FieldID = T0.FieldID and T2.LanID = " & Session("LanID") & " " & _
		  		  "left outer join OLKCUFDGroups T3 on T3.TableID = T1.TableID and T3.GroupID = IsNull(T1.GroupID, -1) " & _
				  "where T0.TableId = 'ORCT' and AType in ('V','T') and Active = 'Y' " & _
				  "order by T3.[Order], IsNull(T1.Pos, 'D'), IsNull(T1.[Order], 32727) "
			rsUFD.open sql, conn, 3, 1
			If rsUFD.recordcount > 0 Then
			sql = "select "
			do while not rsUFD.eof
				If rsUFD.bookmark <> 1 Then sql = sql & ", "
				If rsUFD("DropDown") = "Y" Then
					sql = sql & "(select IsNull((select AlterDescr from OLKUFD1AlterNames where TableID = UFD1.TableID and FieldID = UFD1.FieldID and IndexID = UFD1.IndexID and LanID = " & Session("LanID") & "), Descr) from UFD1 where FldValue = T0." & rsUFD("InsertID") & " collate database_default and TableID = 'ORCT' and FieldID = " & rsUFD("FieldID") & ") " & rsUFD("InsertID")
				Else
					sql = sql & rsUFD("InsertID") & " As '" & rsUFD("InsertID") & "'"
				End If
			rsUFD.movenext
			loop
			sql = sql & " from " & TableID & " T0 where " & sqlDocNum & " = " & Request("DocEntry")
			set rcVal = Server.CreateObject("ADODB.RecordSet")
			set rcVal = conn.execute(sql)
			rsUFD.movefirst
			End If

           set rs = Server.CreateObject("ADODB.recordset")
           
           If Request("DocType") = "-2" or docType = 24 and rc("DocType") = "C" or docType = 46 and rc("DocType") = "S" Then
           	set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKCXCRctDetails" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@ObjectCode") = CInt(Request("DocType"))
			If docType <> "" Then cmd("@DocType") = docType
			cmd("@DocEntry") = Request("DocEntry")
			cmd("@MainCur") = myApp.MainCur
			set rs = cmd.execute()

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
				strContent = Replace(strContent, "{CmpName}", CmpName)
%>
<% If Request("Excell") <> "Y" and Request("PDF") <> "Y" Then %>
<script type="text/javascript">
<!--
function saveDoc(cmd)
{
	switch (cmd)
	{
		case 'Print':
			document.getElementById('tblSave').style.display = 'none';
			document.getElementById('crdLink').style.display = 'none';
			showDocLink(false);
			window.print();
			showDocLink(true);
			document.getElementById('tblSave').style.display = '';
			document.getElementById('crdLink').style.display = '';
			break;
		case 'PDF':
			document.frmExcell.action = 'cxcRctDetailPDF.asp';
			document.frmExcell.submit();
			break;
		case 'Excell':
			document.frmExcell.action = 'cxcRctDetailOpen.asp';
			document.frmExcell.submit();
			break;
	}
}

function showDocLink(show)
{
	if (document.getElementById('docLink') != null)
	{
		myLinks = document.getElementById('docLink');
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
<form target="_blank" method="post" name="frmViewDetail" action="">
<input type="hidden" name="DocEntry" value="">
<input type="hidden" name="DocType" value="">
<input type="hidden" name="CardCode" value="">
<input type="hidden" name="pop" value="Y">
</form>
<script language="javascript">
function goDetail(DocType, DocEntry) {
	if (DocType == 2)
	{
		document.frmViewDetail.action = 'addCard/crdConfDetailOpen.asp';
		document.frmViewDetail.CardCode.value = DocEntry;
	}
	else if (DocType != 24)
	{
		document.frmViewDetail.action = "cxcDocDetailOpen.asp";
		document.frmViewDetail.DocEntry.value = DocEntry;
	}
	document.frmViewDetail.DocType.value = DocType;
	document.frmViewDetail.submit();
}
</script>
<% If Request("Excell") <> "Y" and Request("PDF") <> "Y" Then %>
<table border="0" cellpadding="0" cellspacing="0" width="100%" id="tblSave">
	<tr>
		<td align="right">
		<a href="#" onclick="javascript:saveDoc('Print');"><img alt="<%=getcxcRctDetailLngStr("DtxtPrint")%>" border="0" src="images/print_OLK.gif"></a>&nbsp;
		<% If userType = "C" or userType = "V" and myAut.HasAuthorization(65) Then %><a href="#" onclick="javascript:saveDoc('PDF');"><img alt="<%=getcxcRctDetailLngStr("DtxtExpPDF")%>" border="0" src="images/pdf_OLK.gif"></a>&nbsp;<% End If %>
		<% If userType = "C" or userType = "V" and myAut.HasAuthorization(64) Then %><a href="#" onclick="javascript:saveDoc('Excell');"><img alt="<%=getcxcRctDetailLngStr("LtxtExpToExcell")%>" border="0" src="images/excell.gif"></a><% End If %>
		</td>
	</tr>
</table>
<% End If %>

<table border="0" cellpadding="0" width="100%" id="table1">
	<%=strContent%>
	<!--#include file="myTitle.asp"-->
	<% If CInt(Request("DocType")) = 140 Then %>
	<tr>
		<td id="tdMyTtl" class="TablasTituloDraft">
		<%=getcxcRctDetailLngStr("LttlDraftNote")%></td>
	</tr>
	<% End If %>
	<tr class="CanastaTitle">
		<td><b>
		<% If Request("DocType") = "24" or Request("DocType") = "46" Then %><% If rc("DocNum") <> "" Then %><% If 1 = 2 Then %>Recibo<% Else %><% If Request("DocType") = "24" Then %><%=txtRct%><% Else %><%=txtOvpm%><% End If %><% End If %> 
		#<%=rc("DocNum")%>&nbsp;<% End If %>
		<% Else %><%=getcxcRctDetailLngStr("DtxtLogNum")%> #<%=Request("DocEntry")%><% End If %></b></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table2">
			<tr>
				<td valign="top" width="50%">
				<table border="0" cellpadding="0" width="100%" cellspacing="1" id="table3">
					<tr>
						<td class="CanastaTblResaltada" width="121">
						<table cellpadding="0" cellspacing="0" border="0" width="100%">
							<tr class="CanastaTblResaltada">
								<td><%=getcxcRctDetailLngStr("DtxtCode")%></td>
								<% If Request("Excell") <> "Y" and Request("PDF") <> "Y" Then %><td align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
								<a href="javascript:goDetail(2, '<%=rc("CardCode")%>')"><img id="crdLink" src="design/<%=SelDes%>/images/<%=Session("rtl")%>flecha_selec.gif" border="0"></a></td><% End If %>
							</tr>
						</table>
						</td>
						<td class="CanastaTbl"><%=rc("CardCode")%>&nbsp;</td>
					</tr>
					<tr>
						<td class="CanastaTblResaltada" width="121"><%=getcxcRctDetailLngStr("DtxtName")%></td>
						<td class="CanastaTbl"><%=rc("CardName")%>&nbsp;</td>
					</tr>
					<tr>
						<td class="CanastaTblResaltada" width="121">
						<%=getcxcRctDetailLngStr("DtxtAddress")%></td>
						<td class="CanastaTbl"><%=rc("Address")%>&nbsp;</td>
					</tr>
					<tr>
						<td class="CanastaTblResaltada" width="121">
						<%=getcxcRctDetailLngStr("DtxtContact")%></td>
						<td class="CanastaTbl"><%=rc("CntctName")%>&nbsp;</td>
					</tr>

          			<% rsUFD.Filter = "Pos = 'I'"
          			do while not rsUFD.eof
          			rsFieldName = rsUFD("InsertID") %>
					<tr>
                      <td class="CanastaTblResaltada" width="121">
                      <%=rsUFD("Descr")%>&nbsp;</td>
                      <td class="CanastaTbl">
                      <% If rsUFD("TypeID") = "M" and rsUFD("EditType") = "B" Then %><a class="LinkNoticiasMas" target="_blank" href="<%=rcVal(rsFieldName)%>"><% End If %>
                      <% If rsUFD("TypeID") = "B" Then
                      		If Not IsNull(rcVal(rsFieldName)) Then
			            	Select Case rsUFD("EditType")
								Case "R"
									Response.Write FormatNumber(CDbl(rcVal(rsFieldName)),myApp.RateDec)
								Case "S"
									Response.Write FormatNumber(CDbl(rcVal(rsFieldName)),myApp.SumDec)
								Case "P"
									Response.Write FormatNumber(CDbl(rcVal(rsFieldName)),myApp.PriceDec)
								Case "Q"
									Response.Write FormatNumber(CDbl(rcVal(rsFieldName)),myApp.QtyDec)
								Case "%"
									Response.Write FormatNumber(CDbl(rcVal(rsFieldName)),myApp.PercentDec)
								Case "M"
									Response.Write FormatNumber(CDbl(rcVal(rsFieldName)),myApp.MeasureDec)
			            	End Select
			            	End If
                      ElseIf rsUFD("TypeID") = "A" and rsUFD("EditType") = "I" Then %>
                      <img src="pic.aspx?filename=<% If Not IsNull(rcVal(rsFieldName)) Then %><%=rcVal(rsFieldName)%><% Else %>n_a.gif<% End If %>&MaxSize=180&dbName=<%=Session("olkdb")%>" border="0">
                      <% ElseIf rsUFD("TypeID") = "D" Then %>
                      <%=FormatDate(rcVal(rsFieldName), True)%>
                      <% Else %>
                      <% If Not IsNull(rcVal(rsFieldName)) Then %><%=rcVal(rsFieldName)%><% End If %>
                      <% End If %>
                      <% If rsUFD("TypeID") = "M" and rsUFD("EditType") = "B" Then %></a><% End If %>
                      </td>
                    </tr>
          			<% rsUFD.movenext
          			loop
          			if rsUFD.recordcount > 0 then
          				rsUFD.movefirst
          			end if
          			rsUFD.filter = "Pos = 'D'" %>
				</table>
				</td>
				<td valign="top" width="50%">
				<table border="0" cellpadding="0" width="100%" cellspacing="1" id="table4">
					<tr>
						<td class="CanastaTblResaltada" width="128"><%=getcxcRctDetailLngStr("DtxtDate")%></td>
						<td class="CanastaTbl"><%=FormatDate(rc("DocDate"), True)%>&nbsp;</td>
					</tr>
					<tr>
						<td class="CanastaTblResaltada" width="128">
						<%=getcxcRctDetailLngStr("LtxtCounterRef")%></td>
						<td class="CanastaTbl"><%=rc("CounterRef")%>&nbsp;</td>
					</tr>
					<% do while not rsUFD.eof
					rsFieldName = rsUFD("InsertID") %>
          			<tr>
                      <td class="CanastaTblResaltada" width="128">
                      <%=rsUFD("Descr")%>&nbsp;</td>
                      <td class="CanastaTbl">
                      <% If rsUFD("TypeID") = "M" and rsUFD("EditType") = "B" Then %><a class="LinkNoticiasMas" target="_blank" href="<%=rcVal(rsFieldName)%>"><% End If %>
                      <% If rsUFD("TypeID") = "B" Then
                      		If Not IsNull(rcVal(rsFieldName)) Then
			            	Select Case rsUFD("EditType")
								Case "R"
									Response.Write FormatNumber(CDbl(rcVal(rsFieldName)),myApp.RateDec)
								Case "S"
									Response.Write FormatNumber(CDbl(rcVal(rsFieldName)),myApp.SumDec)
								Case "P"
									Response.Write FormatNumber(CDbl(rcVal(rsFieldName)),myApp.PriceDec)
								Case "Q"
									Response.Write FormatNumber(CDbl(rcVal(rsFieldName)),myApp.QtyDec)
								Case "%"
									Response.Write FormatNumber(CDbl(rcVal(rsFieldName)),myApp.PercentDec)
								Case "M"
									Response.Write FormatNumber(CDbl(rcVal(rsFieldName)),myApp.MeasureDec)
			            	End Select
			            	End If
                      ElseIf rsUFD("TypeID") = "A" and rsUFD("EditType") = "I" Then %>
                      <img src="pic.aspx?filename=<% If Not IsNull(rcVal(rsFieldName)) Then %><%=rcVal(rsFieldName)%><% Else %>n_a.gif<% End If %>&MaxSize=180&dbName=<%=Session("olkdb")%>" border="0">
                      <% ElseIf rsUFD("TypeID") = "D" Then %>
                      <%=FormatDate(rcVal(rsFieldName), True)%>
                      <% Else %>
                      <% If Not IsNull(rcVal(rsFieldName)) Then %><%=rcVal(rsFieldName)%><% End If %>
                      <% End If %>
                      <% If rsUFD("TypeID") = "M" and rsUFD("EditType") = "B" Then %></a><% End If %>
                      </td>
                    </tr>
          			<% rsUFD.movenext
          			loop %>
				</table>
				</td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<% If Request("DocType") = "-2" or docType = 24 and rc("DocType") = "C" or docType = 46 and rc("DocType") = "S" Then %>
			<tr class="FirmTlt3">
				<td>
				<p align="center">#</td>
				<td align="center"><%=getcxcRctDetailLngStr("DtxtInstallment")%></td>
				<td align="center"><%=getcxcRctDetailLngStr("DtxtDate")%></td>
				<td align="center"><%=getcxcRctDetailLngStr("DtxtDetail")%></td>
				<td align="center"><%=getcxcRctDetailLngStr("LtxtDocTotal")%></td>
				<td align="center" width="143"><%=getcxcRctDetailLngStr("DtxtPaid")%></td>
				<td align="center" width="128"><%=getcxcRctDetailLngStr("DtxtBalance")%></td>
			</tr>
		      <% TotalSaldo = 0
		      do while not rs.eof %>
			<tr class="CanastaTbl">
				<td><% If Request("Excell") <> "Y" and Request("PDF") <> "Y" Then %>
				<a href="javascript:goDetail(<%=rs("ObjType")%>, '<%=rs("DocEntry")%>')"><img id="docLink" src="design/<%=SelDes%>/images/<%=Session("rtl")%>flecha_selec.gif" border="0"></a><% End If %><%=rs("DocNum")%></td>
				<td><%=Replace(Replace(getcxcRctDetailLngStr("DtxtXofY"), "{0}", rs("InstID")), "{1}", rs("InstCount"))%></td>
				<td align="center"><%=rs("DocDate")%>&nbsp;</td>
				<td align="center"><%=rs("Comments")%>&nbsp;</td>
				<td align="right"><nobr><%=rs("DocCur")%>&nbsp;<%=FormatNumber(rs("DocTotal"),myApp.SumDec)%></nobr></td>
				<td align="center" width="143">
				<p align="right"><nobr><%=rs("DocCur")%>&nbsp;<%=FormatNumber(rs("SumApplied"),myApp.SumDec)%></nobr></td>
				<td align="right" width="128"><nobr><%=rs("DocCur")%>&nbsp;<%=FormatNumber(rs("Saldo"),myApp.SumDec)%></nobr></td>
			</tr>
		      <% rs.movenext
		      loop %>
		    <% End If %>
			<tr>
				<td align="center" colspan="5" rowspan="3" valign="top">
				<table border="0" cellpadding="0" width="100%" id="table6">
					<tr class="FirmTlt3">
						<td align="center"><%=getcxcRctDetailLngStr("LtxtCash")%></td>
						<td align="center"><%=getcxcRctDetailLngStr("LtxtChecks")%></td>
						<td width="128" align="center"><%=getcxcRctDetailLngStr("LtxtCredCard")%></td>
						<td width="169" align="center"><%=getcxcRctDetailLngStr("LtxtBankTrans")%></td>
					</tr>
					<tr class="CanastaTbl">
						<td align="center"><nobr><%=rc("DocCur")%>&nbsp;<%=FormatNumber(Rc("CashSum"),myApp.SumDec)%></nobr></td>
						<td align="center"><nobr><%=rc("DocCur")%>&nbsp;<%=FormatNumber(Rc("CheckSum"),myApp.SumDec)%></nobr></td>
						<td width="128" align="center"><nobr><%=rc("DocCur")%>&nbsp;<%=FormatNumber(Rc("CreditSum"),myApp.SumDec)%></nobr></td>
						<td width="169" align="center"><nobr><%=rc("DocCur")%>&nbsp;<%=FormatNumber(Rc("TrsfrSum"),myApp.SumDec)%></nobr></td>
					</tr>
				</table>
				</td>
				<td class="CanastaTblResaltada" align="center" width="143">
				<p align="<% If Session("rtl") = "" Then %>left<% Else %>right<% End If %>">
				<%=getcxcRctDetailLngStr("LtxtTotalPaid")%></td>
				<td class="CanastaTbl" align="right" width="128"><nobr><%=rc("DocCur")%>&nbsp;<%=FormatNumber(rc("Pagado"),myApp.SumDec)%></nobr></td>
			</tr>
			<tr>
				<td class="CanastaTblResaltada" align="<% If Session("rtl") = "" Then %>left<% Else %>right<% End If %>" width="143">
				<%=getcxcRctDetailLngStr("LtxtTotalCanceled")%></td>
				<td class="CanastaTbl" align="right" width="128"><nobr><%=rc("DocCur")%>&nbsp;<%=FormatNumber(rc("DocTotal"),myApp.SumDec)%></nobr></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>
		<table width="50%">
			<tr>
				<td class="CanastaTblResaltada" style="width: 100px" valign="top"><nobr><%=getcxcRctDetailLngStr("DtxtNote")%></nobr></td>
				<td class="CanastaTbl"><% If Not IsNull(RC("Comments")) Then %><%=myHTMLEncode(RC("Comments"))%><% End If %></td>
			</tr>
			<tr>
				<td class="CanastaTblResaltada" style="width: 100px" valign="top"><nobr><%=getcxcRctDetailLngStr("DtxtObservations")%></nobr></td>
				<td class="CanastaTbl"><% If Not IsNull(RC("JrnlMemo")) Then %><%=myHTMLEncode(RC("JrnlMemo"))%><% End If %></td>
			</tr>
		</table>
		</td>
	</tr>
	</table>
<% If Request("Excell") <> "Y" and Request("PDF") <> "Y" Then %>
<form name="frmExcell" method="post">
<% For each itm in Request.Form %><input type="hidden" name="<%=itm%>" value="<%=Request.Form(itm)%>"><% Next %>
<input type="hidden" name="Excell" value="Y">
</form>
<% End If %><% set rc = nothing 
set rs = nothing 
conn.close %>


<% Else %>
	<script type="text/javascript">
	alert('<%=getcxcRctDetailLngStr("DtxtNoAccessObj")%>'.replace('{0}', '<% If Request("DocType") = "46" Then %><%=txtOvpm%><% Else %><%=txtRct%><% End If %>'));
	window.close();
	</script>
<% End If %>
<% Else %>
	<script type="text/javascript">
	alert('<%=getcxcRctDetailLngStr("DtxtNoData")%>');
	window.close();
	</script>
<% End If %>
</body>
</html>
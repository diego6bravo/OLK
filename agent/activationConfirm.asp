<!--#include file="clientInc.asp"-->
<!--#include file="agentTop.asp"-->
<% If userType <> "V" Then Response.Redirect "default.asp" %>
<% 
If Not (EnableClientActivation and myAut.HasAuthorization(89)) Then Response.Redirect "unauthorized.asp"
addLngPathStr = "" %>
<!--#include file="lang/activationConfirm.asp" -->
<%
set rs = Server.CreateObject("ADODB.recordset")
%>
<script language="javascript">
function GoLogView(CardCode) {
document.viewLogNum.CardCode.value = CardCode 
document.viewLogNum.submit() }

</script>
<form target="_blank" method="post" name="viewLogNum" action="addCard/crdConfDetailOpen.asp">
<input type="hidden" name="CardCode" value="">
<input type="hidden" name="DocType" value="2">
<input type="hidden" name="pop" value="Y">
<input type="hidden" name="AddPath" value="../">
</form>
<form method="POST" action="activationConfirm.asp" name="frm">
<table border="0" cellpadding="0" width="100%">
	<tr class="GeneralTlt">
		<td>&nbsp;<%=getactivationConfirmLngStr("LttlAnonRegAct")%></td>
	</tr>
	<tr class="GeneralTblBold2">
		<td><% If 1 = 2 Then %>Clientes<% Else %><%=txtClients%><% End If %></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table6">
			<tr class="GeneralTblBold2">
				<td align="center" width="15">&nbsp;</td>
				<td align="center"><%=getactivationConfirmLngStr("DtxtCode")%></td>
				<td align="center"><%=getactivationConfirmLngStr("DtxtName")%></td>
				<td align="center"><%=getactivationConfirmLngStr("DtxtGroup")%></td>
				<td align="center"><%=getactivationConfirmLngStr("DtxtCountry")%></td>
				<td align="center"><%=getactivationConfirmLngStr("DtxtDate")%></td>
				<td align="center"><%=getactivationConfirmLngStr("DtxtState")%></td>
				<td align="center"><% If 1 = 2 Then %>|X:AsignAgent|<% Else %><%=Replace(getactivationConfirmLngStr("LtxtAsignedAgent"), "{0}", txtAgent)%><% End If %></td>
				<td align="center">&nbsp;</td>
			</tr>
			<%
			uSql = "declare @mailID int "
			set rCrd = Server.CreateObject("ADODB.RecordSet")
			If Session("ConfRetVal") <> "" Then
				sqlAddStr = ", L0.Status R3Status, L0.ErrMessage, L0.LogNum "
				innerStr = 	"left outer join R3_ObsCommon..TLOG L0 on L0.Company = db_name() and L0.ObjectCode = T0.CardCode collate database_default " & _
							"and L0.LogNum = (select Max(LogNum) from R3_ObsCommon..TLOG where Company = db_name() and ObjectCode = T0.CardCode collate database_default) "
			Else
				sqlAddStr = ", null R3Status, null ErrMessage "
			End If
			sql = "select T0.CardCode, T0.DocEntry, IsNull(T0.CardName, '') CardName, IsNull(T1.GroupName, '') GroupName, IsNull(T2.Name, '') Country, T0.CreateDate, T0.DocEntry, IsNull(SlpName, '') SlpName, " & _
			"T3.Status, Case T3.Status When 'A' Then '" & getactivationConfirmLngStr("DtxtActive") & "' When 'R' Then '" & getactivationConfirmLngStr("DtxtRejected") & "' End StatusStr" & sqlAddStr & " " & _
			"from OCRD T0 " & _
			"inner join OCRG T1 on T1.GroupCode = T0.GroupCode " & _
			"left outer join OCRY T2 on T2.Code = T0.Country " & _
			"inner join OLKClientsAccess T3 on T3.CardCode = T0.CardCode " & _
			"inner join OSLP T4 on T4.SlpCode = T0.SlpCode " & innerStr & _
			"where T0.DocEntry in (" & Request("ConfDocEntry") & ") order by T0.CardCode asc "
			set rCrd  = conn.execute(sql)
			'response.write sql
			if not rCrd.eof then
			do while not rCrd.eof
			If Session("NotifyAdd") Then
				If rCrd("R3Status") = "S" or IsNull(rCrd("R3Status")) and rCrd("Status") = "R" Then
					If rCrd("R3Status") = "S" Then TypeID = 1 Else MailID = 3
				End If
				If rCrd("R3Status") = "S" Then
					uSql = uSql & "update OLKClientsAccess set Status = 'A', StatusDate = getdate(), " & _
								"StatusUserSign = " & Session("vendid") & " " & _
								"where CardCode = N'" & rCrd("CardCode") & "' " & _
								"set @mailID = IsNull((select Max(mailID)+1 from OLKMail), 0) " & _
								"insert OLKMail(mailID, TypeID, LanID, Entry, Sent) " & _
								"values(@mailID, 1, " & Session("LanID") & ", " & rCrd("LogNum") & ", 'N') "
				ElseIf IsNull(rCrd("R3Status")) and rCrd("Status") = "R" Then
					uSql = uSql & "set @mailID = IsNull((select Max(mailID)+1 from OLKMail), 0) " & _
								"insert OLKMail(mailID, TypeID, LanID, Entry, Sent) " & _
								"values(@mailID, 3, " & Session("LanID") & ", " & rCrd("DocEntry") & ", 'N') "
				End If
			End If
			R3Status = rCrd("R3status") %>
			<tr class="GeneralTbl">
				<td width="15">
				<p align="center"><a href="javascript:GoLogView('<%=myHTMLEncode(rCrd("CardCode"))%>')">
				<img border="0" src="design/0/images/<%=Session("rtl")%>felcahSelect.gif" width="15" height="13"></a>
				</td>
				<td><%=rCrd("CardCode")%>&nbsp;</td>
				<td><%=rCrd("CardName")%>&nbsp;</td>
				<td><%=rCrd("GroupName")%>&nbsp;</td>
				<td><%=rCrd("Country")%>&nbsp;</td>
				<td><%=FormatDate(rCrd("CreateDate"), True)%>&nbsp;</td>
				<td>
				<p align="center"><% If R3Status = "E" Then %><input type="hidden" id="errMsg" name="errMsg<%=rCrd("DocEntry")%>" value="<%=myHTMLEncode(rCrd("errMessage"))%>"><a href="#" class="LinkGeneral" onclick="javascript:alert('<%=getactivationConfirmLngStr("LtxtOBServerErr")%>: \n' + document.frm.errMsg<%=rCrd("DocEntry")%>.value);"><% End If %>
				<% Select Case R3Status
				Case "S"
					Response.write getactivationConfirmLngStr("LtxtActivated")
				Case "E"
					Response.write getactivationConfirmLngStr("DtxtError")
				End Select	
				%><% If R3Status = "E" Then %></a><% End If %>
				<% If IsNull(R3Status) Then %><%=rCrd("StatusStr")%><% End If %></td>
				<td>
				<%=rCrd("SlpName")%>&nbsp;</td>
				<td>
				&nbsp;</td>
			</tr>
			<% rCrd.movenext
			loop %>
			<% Else %>
			<tr class="GeneralTblBold2">
				<td colspan="9">
				<p align="center"><% If 1 = 2 Then %>|X:NoData|<% Else %><%=Replace(getactivationConfirmLngStr("LtxtNoClientActive"), "{0}", txtClient)%><% End If %></td>
			</tr>
			<% End If %>
		</table>
		</td>
	</tr>
	<tr class="GeneralTbl">
		<td>
		<p align="right">
		&nbsp;</td>
	</tr>
  <% sql = "select IsNull(slpname, '') slpname from oslp where slpcode = " & Session("Vendid")
  set rf = conn.execute(sql) %>
	<tr class="GeneralTblBold2">
		<td><%=getactivationConfirmLngStr("LtxtConfBy")%>: <%=rf(0)%></td>
	</tr>
</table>
<input type="hidden" name="cmd" value="confirmdocsSubmit">
</form>
<% If uSql <> "" and Session("NotifyAdd") Then
	conn.execute(uSql)
	Session("NotifyAdd") = False
End If %>
<% 
Function getActivationConfAcceptBody(ByVal txtMsg)

	set rm = Server.CreateObject("ADODB.RecordSet")
	sql = 	"select top 1 IsNull(PrintHeadr, IsNull(CompnyName, '')) PrintHeadr, IsNull(CompnyAddr, '') CompnyAddr, IsNull(Country, '') Country, Phone1, Phone2, Fax, E_Mail, " & _
			"(select SelDes from OLKCommon) SelDes from oadm order by CurrPeriod desc"
	set rm = conn.execute(sql)
	
	cmpMailStr = "<font face=""Verdana"" size=""1"">" & myHTMLEncode(rm("CompnyAddr")) & "<br>"
	If rm("Phone1") <> "" Then cmpMailStr = cmpMailStr & DtxtPhone & ": " & rm("Phone1")
	If rm("Phone2") <> "" Then cmpMailStr = cmpMailStr & "/" & rm("Phone2")
	If rm("Fax") <> "" Then cmpMailStr = cmpMailStr & " " & DtxtFax & ": " & rm("Fax") & "<br>"
	If rm("E_Mail") <> "" Then cmpMailStr = cmpMailStr & "<a class=""LinkTop"" href=""mailto:" & rm("E_Mail") & """>" & _
							rm("E_Mail") & "</a>"
	cmpMailStr = cmpMailStr & "</font>"
	
	imgPath = 	GetHTTPStr & Request.ServerVariables("HTTP_HOST") & Replace(LCase(Request.ServerVariables("URL")),"activationconfirm.asp","") & _
				"design/" & rm("SelDes") & "/mail/2/images/"
				
	If myApp.MailLogo or IsNull(myApp.MailLogo) Then 
		MailLogo = imgPath & "validacion_registro_r2_c4.jpg" 
	Else 
		MailLogo = GetHTTPStr & Request.ServerVariables("HTTP_HOST") & Replace(LCase(Request.ServerVariables("URL")),"activationconfirm.asp","") & "imagenes/" & Session("olkdb") & "/" & myApp.MailLogo
	End If
	
	msgStr = "<html> " & _
	"<head> " & _
	"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8""> " & _
	"<title>OLK - Cuenta Activada</title> " & _
	"</head> " & _
	"<body style=""text-align: center"" topmargin=""0""> " & _
	"<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""538""> " & _
	"  	<tr> " & _
	"		<td> " & _
	"		<p align=""center""> " & _
	"		<img src=""" & imgPath & "spacer.gif"" width=""13"" height=""1"" border=""0"" alt=""""></td> " & _
	"		<td> " & _
	"		<p align=""center""> " & _
	"		<img src=""" & imgPath & "spacer.gif"" width=""141"" height=""1"" border=""0"" alt=""""></td> " & _
	"		<td> " & _
	"		<p align=""center""> " & _
	"		<img src=""" & imgPath & "spacer.gif"" width=""125"" height=""1"" border=""0"" alt=""""></td> " & _
	"		<td> " & _
	"		<p align=""center""> " & _
	"		<img src=""" & imgPath & "spacer.gif"" width=""242"" height=""1"" border=""0"" alt=""""></td> " & _
	"		<td> " & _
	"		<p align=""center""> " & _
	"		<img src=""" & imgPath & "spacer.gif"" width=""5"" height=""1"" border=""0"" alt=""""></td> " & _
	"		<td> " & _
	"		<p align=""center""> " & _
	"		<img src=""" & imgPath & "spacer.gif"" width=""12"" height=""1"" border=""0"" alt=""""></td> " & _
	"		<td> " & _
	"		<p align=""center""> " & _
	"		<img src=""" & imgPath & "spacer.gif"" width=""1"" height=""1"" border=""0"" alt=""""></td> " & _
	"	</tr> " & _
	"	<tr> " & _
	"		<td colspan=""6""> " & _
	"		<p align=""center""> " & _
	"		<img name=""validacion_registro_r1_c1"" src=""" & imgPath & "validacion_registro_r1_c1.jpg"" width=""538"" height=""25"" border=""0"" alt=""""></td> " & _
	"		<td> " & _
	"		<p align=""center""> " & _
	"		<img src=""" & imgPath & "spacer.gif"" width=""1"" height=""25"" border=""0"" alt=""""></td> " & _
	"	</tr> " & _
	"	<tr> " & _
	"		<td colspan=""3""> " & _
	"		<p align=""center""> " & _
	"		<img name=""validacion_registro_r2_c1"" src=""" & imgPath & "validacion_registro_r2_c1.jpg"" width=""279"" height=""164"" border=""0"" alt=""""></td> " & _
	"		<td colspan=""2""> " & _
	"		<p align=""center""> " & _
	"		<img name=""validacion_registro_r2_c4"" src=""" & MailLogo & """ width=""247"" height=""164"" border=""0"" alt=""""></td> " & _
	"		<td> " & _
	"		<p align=""center""> " & _
	"		<img name=""validacion_registro_r2_c6"" src=""" & imgPath & "validacion_registro_r2_c6.jpg"" width=""12"" height=""164"" border=""0"" alt=""""></td> " & _
	"		<td> " & _
	"		<p align=""center""> " & _
	"		<img src=""" & imgPath & "spacer.gif"" width=""1"" height=""164"" border=""0"" alt=""""></td> " & _
	"	</tr> " & _
	"	<tr> " & _
	"		<td colspan=""6""> " & _
	"		<p align=""center""> " & _
	"		<img name=""validacion_registro_r3_c1"" src=""" & imgPath & "validacion_registro_r3_c1.jpg"" width=""538"" height=""41"" border=""0"" alt=""""></td> " & _
	"		<td> " & _
	"		<p align=""center""> " & _
	"		<img src=""" & imgPath & "spacer.gif"" width=""1"" height=""41"" border=""0"" alt=""""></td> " & _
	"	</tr> " & _
	"	<tr> " & _
	"		<td colspan=""6""> " & _
	"		<p align=""center""> " & _
	"		<img name=""validacion_registro_r4_c1"" src=""" & imgPath & "validacion_registro_r4_c1.jpg"" width=""538"" height=""13"" border=""0"" alt=""""></td> " & _
	"		<td> " & _
	"		<p align=""center""> " & _
	"		<img src=""" & imgPath & "spacer.gif"" width=""1"" height=""13"" border=""0"" alt=""""></td> " & _
	"	</tr> " & _
	"	<tr> " & _
	"		<td background=""" & imgPath & "validacion_registro_r5_c1.jpg""> " & _
	"		<p align=""center""> " & _
	"		<img name=""validacion_registro_r5_c1"" src=""" & imgPath & "validacion_registro_r5_c1.jpg"" width=""13"" height=""241"" border=""0"" alt=""""></td> " & _
	"		<td colspan=""3"" background=""" & imgPath & "background.jpg"" valign=""top""> " & _
	"		<table border=""0"" cellpadding=""0"" width=""100%"" id=""table3""> " & _
	"			<tr> " & _
	"				<td height=""80"">&nbsp;</td> " & _
	"			</tr> " & _
	"			<tr> " & _
	"				<td> " & _
	"				<p align=""center""><font face=""Verdana"" size=""2"">" & txtMsg & "</font></td> " & _
	"			</tr> " & _
	"		</table> " & _
	"		</td> " & _
	"		<td colspan=""2"" background=""" & imgPath & "validacion_registro_r5_c5.jpg""> " & _
	"		<p align=""center""> " & _
	"		<img name=""validacion_registro_r5_c5"" src=""" & imgPath & "validacion_registro_r5_c5.jpg"" width=""17"" height=""241"" border=""0"" alt=""""></td> " & _
	"		<td> " & _
	"		<p align=""center""> " & _
	"		<img src=""" & imgPath & "spacer.gif"" width=""1"" height=""241"" border=""0"" alt=""""></td> " & _
	"	</tr> " & _
	"	<tr> " & _
	"		<td colspan=""6""> " & _
	"		<p align=""center""> " & _
	"		<img name=""validacion_registro_r6_c1"" src=""" & imgPath & "validacion_registro_r6_c1.jpg"" width=""538"" height=""12"" border=""0"" alt=""""></td> " & _
	"		<td> " & _
	"		<p align=""center""> " & _
	"		<img src=""" & imgPath & "spacer.gif"" width=""1"" height=""12"" border=""0"" alt=""""></td> " & _
	"	</tr> " & _
	"	<tr> " & _
	"		<td rowspan=""2""> " & _
	"		<p align=""center""> " & _
	"		<img name=""validacion_registro_r7_c1"" src=""" & imgPath & "validacion_registro_r7_c1.jpg"" width=""13"" height=""92"" border=""0"" alt=""""></td> " & _
	"		<td> " & _
	"		<p align=""center""> " & _
	"		<img name=""validacion_registro_r7_c2"" src=""" & imgPath & "validacion_registro_r7_c2.jpg"" width=""141"" height=""75"" border=""0"" alt=""""></td> " & _
	"		<td colspan=""2"" background=""" & imgPath & "validacion_registro_r7_c3.jpg"" valign=""top""> " & _
	"		<table border=""0"" cellpadding=""0"" width=""100%"" cellspacing=""1"" id=""table4""> " & _
	"			<tr> " & _
	"				<td>&nbsp;</td> " & _
	"				<td>&nbsp;</td> " & _
	"			</tr> " & _
	"			<tr> " & _
	"				<td> " & _
	"				<p align=""right"">" & cmpMailStr & "</td> " & _
	"				<td> " & _
	"				&nbsp;</td> " & _
	"			</tr> " & _
	"		</table> " & _
	"		</td> " & _
	"		<td rowspan=""2"" colspan=""2""> " & _
	"		<p align=""center""> " & _
	"		<img name=""validacion_registro_r7_c5"" src=""" & imgPath & "validacion_registro_r7_c5.jpg"" width=""17"" height=""92"" border=""0"" alt=""""></td> " & _
	"		<td> " & _
	"		<p align=""center""> " & _
	"		<img src=""" & imgPath & "spacer.gif"" width=""1"" height=""75"" border=""0"" alt=""""></td> " & _
	"	</tr> " & _
	"	<tr> " & _
	"		<td colspan=""3""> " & _
	"		<p align=""center""> " & _
	"		<img name=""validacion_registro_r8_c2"" src=""" & imgPath & "validacion_registro_r8_c2.jpg"" width=""508"" height=""17"" border=""0"" alt=""""></td> " & _
	"		<td> " & _
	"		<p align=""center""> " & _
	"		<img src=""" & imgPath & "spacer.gif"" width=""1"" height=""17"" border=""0"" alt=""""></td> " & _
	"	</tr> " & _
	"</table> " & _
	"</body> " & _
	"</html> "
	
	getActivationConfAcceptBody = msgStr
End Function 

Function getActivationConfRejBody(ByVal txtMsg)

	set rm = Server.CreateObject("ADODB.RecordSet")
	sql = 	"select top 1 IsNull(PrintHeadr, CompnyName) PrintHeadr, CompnyAddr, Country, Phone1, Phone2, Fax, E_Mail, " & _
			"(select SelDes from OLKCommon) SelDes from oadm order by CurrPeriod desc"
	set rm = conn.execute(sql)
	
	cmpMailStr = "<font face=""Verdana"" size=""1"">" & myHTMLEncode(rm("CompnyAddr")) & "<br>"
	If rm("Phone1") <> "" Then cmpMailStr = cmpMailStr & DtxtPhone & ": " & rm("Phone1")
	If rm("Phone2") <> "" Then cmpMailStr = cmpMailStr & "/" & rm("Phone2")
	If rm("Fax") <> "" Then cmpMailStr = cmpMailStr & " " & txtFax & ": " & rm("Fax") & "<br>"
	If rm("E_Mail") <> "" Then cmpMailStr = cmpMailStr & "<a class=""LinkTop"" href=""mailto:" & rm("E_Mail") & """>" & _
							rm("E_Mail") & "</a>"
	cmpMailStr = cmpMailStr & "</font>"
	
	imgPath = 	GetHTTPStr & Request.ServerVariables("HTTP_HOST") & Replace(LCase(Request.ServerVariables("URL")),"activationconfirm.asp","") & _
				"design/" & rm("SelDes") & "/mail/5/images/"
				
	If myApp.MailLogo = "" or IsNull(myApp.MailLogo) Then 
		MailLogo = imgPath & "rechazocuenta_r2_c4.jpg" 
	Else 
		MailLogo = GetHTTPStr & Request.ServerVariables("HTTP_HOST") & Replace(LCase(Request.ServerVariables("URL")),"activationconfirm.asp","") & "imagenes/" & Session("olkdb") & "/" & myApp.MailLogo
	End If
	
	msgStr = "<html> " & _
	"<head> " & _
	"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8""> " & _
	"<title>New Page 1</title> " & _
	"</head> " & _
	"<body style=""text-align: center"" topmargin=""0""> " & _
	"<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""538""> " & _
	"  	<tr> " & _
	"		<td> " & _
	"		<p align=""center""> " & _
	"		<img src=""" & imgPath & "spacer.gif"" width=""13"" height=""1"" border=""0"" alt=""""></td> " & _
	"		<td> " & _
	"		<p align=""center""> " & _
	"		<img src=""" & imgPath & "spacer.gif"" width=""141"" height=""1"" border=""0"" alt=""""></td> " & _
	"		<td> " & _
	"		<p align=""center""> " & _
	"		<img src=""" & imgPath & "spacer.gif"" width=""125"" height=""1"" border=""0"" alt=""""></td> " & _
	"		<td> " & _
	"		<p align=""center""> " & _
	"		<img src=""" & imgPath & "spacer.gif"" width=""242"" height=""1"" border=""0"" alt=""""></td> " & _
	"		<td> " & _
	"		<p align=""center""> " & _
	"		<img src=""" & imgPath & "spacer.gif"" width=""5"" height=""1"" border=""0"" alt=""""></td> " & _
	"		<td> " & _
	"		<p align=""center""> " & _
	"		<img src=""" & imgPath & "spacer.gif"" width=""12"" height=""1"" border=""0"" alt=""""></td> " & _
	"		<td> " & _
	"		<p align=""center""> " & _
	"		<img src=""" & imgPath & "spacer.gif"" width=""1"" height=""1"" border=""0"" alt=""""></td> " & _
	"	</tr> " & _
	"	<tr> " & _
	"		<td colspan=""6""> " & _
	"		<p align=""center""> " & _
	"		<img name=""rechazocuenta_r1_c1"" src=""" & imgPath & "rechazocuenta_r1_c1.jpg"" width=""538"" height=""25"" border=""0"" alt=""""></td> " & _
	"		<td> " & _
	"		<p align=""center""> " & _
	"		<img src=""" & imgPath & "spacer.gif"" width=""1"" height=""25"" border=""0"" alt=""""></td> " & _
	"	</tr> " & _
	"	<tr> " & _
	"		<td colspan=""3""> " & _
	"		<p align=""center""> " & _
	"		<img name=""rechazocuenta_r2_c1"" src=""" & imgPath & "rechazocuenta_r2_c1.jpg"" width=""279"" height=""164"" border=""0"" alt=""""></td> " & _
	"		<td colspan=""2""> " & _
	"		<p align=""center""> " & _
	"		<img name=""rechazocuenta_r2_c4"" src=""" & MailLogo & """ width=""247"" height=""164"" border=""0"" alt=""""></td> " & _
	"		<td> " & _
	"		<p align=""center""> " & _
	"		<img name=""rechazocuenta_r2_c6"" src=""" & imgPath & "rechazocuenta_r2_c6.jpg"" width=""12"" height=""164"" border=""0"" alt=""""></td> " & _
	"		<td> " & _
	"		<p align=""center""> " & _
	"		<img src=""" & imgPath & "spacer.gif"" width=""1"" height=""164"" border=""0"" alt=""""></td> " & _
	"	</tr> " & _
	"	<tr> " & _
	"		<td colspan=""6""> " & _
	"		<p align=""center""> " & _
	"		<img name=""rechazocuenta_r3_c1"" src=""" & imgPath & "rechazocuenta_r3_c1.jpg"" width=""538"" height=""41"" border=""0"" alt=""""></td> " & _
	"		<td> " & _
	"		<p align=""center""> " & _
	"		<img src=""" & imgPath & "spacer.gif"" width=""1"" height=""41"" border=""0"" alt=""""></td> " & _
	"	</tr> " & _
	"	<tr> " & _
	"		<td colspan=""6""> " & _
	"		<p align=""center""> " & _
	"		<img name=""rechazocuenta_r4_c1"" src=""" & imgPath & "rechazocuenta_r4_c1.jpg"" width=""538"" height=""13"" border=""0"" alt=""""></td> " & _
	"		<td> " & _
	"		<p align=""center""> " & _
	"		<img src=""" & imgPath & "spacer.gif"" width=""1"" height=""13"" border=""0"" alt=""""></td> " & _
	"	</tr> " & _
	"	<tr> " & _
	"		<td background=""" & imgPath & "rechazocuenta_r5_c1.jpg"" valign=""top""> " & _
	"		<p align=""center""> " & _
	"		<img name=""rechazocuenta_r5_c1"" src=""" & imgPath & "rechazocuenta_r5_c1.jpg"" width=""13"" height=""241"" border=""0"" alt=""""></td> " & _
	"		<td colspan=""3"" background=""" & imgPath & "rechazocuenta_r5_c2.jpg"" valign=""top""> " & _
	"		<table border=""0"" cellpadding=""0"" width=""100%"" id=""table1""> " & _
	"			<tr> " & _
	"				<td height=""80"">&nbsp;</td> " & _
	"			</tr> " & _
	"			<tr> " & _
	"				<td> " & _
	"				<p align=""center""><font face=""Verdana"" size=""2"">" & txtMsg & "</font></td> " & _
	"			</tr> " & _
	"		</table> " & _
	"		</td> " & _
	"		<td colspan=""2"" background=""" & imgPath & "rechazocuenta_r5_c5.jpg"" valign=""top""> " & _
	"		<p align=""center""> " & _
	"		<img name=""rechazocuenta_r5_c5"" src=""" & imgPath & "rechazocuenta_r5_c5.jpg"" width=""17"" height=""241"" border=""0"" alt=""""></td> " & _
	"		<td> " & _
	"		<p align=""center""> " & _
	"		<img src=""" & imgPath & "spacer.gif"" width=""1"" height=""241"" border=""0"" alt=""""></td> " & _
	"	</tr> " & _
	"	<tr> " & _
	"		<td colspan=""6""> " & _
	"		<p align=""center""> " & _
	"		<img name=""rechazocuenta_r6_c1"" src=""" & imgPath & "rechazocuenta_r6_c1.jpg"" width=""538"" height=""12"" border=""0"" alt=""""></td> " & _
	"		<td> " & _
	"		<p align=""center""> " & _
	"		<img src=""" & imgPath & "spacer.gif"" width=""1"" height=""12"" border=""0"" alt=""""></td> " & _
	"	</tr> " & _
	"	<tr> " & _
	"		<td rowspan=""2""> " & _
	"		<p align=""center""> " & _
	"		<img name=""rechazocuenta_r7_c1"" src=""" & imgPath & "rechazocuenta_r7_c1.jpg"" width=""13"" height=""92"" border=""0"" alt=""""></td> " & _
	"		<td> " & _
	"		<p align=""center""> " & _
	"		<img name=""rechazocuenta_r7_c2"" src=""" & imgPath & "rechazocuenta_r7_c2.jpg"" width=""141"" height=""75"" border=""0"" alt=""""></td> " & _
	"		<td colspan=""2"" background=""" & imgPath & "rechazocuenta_r7_c3.jpg"" valign=""top""> " & _
	"		<table border=""0"" cellpadding=""0"" width=""100%"" cellspacing=""1"" id=""table2""> " & _
	"			<tr> " & _
	"				<td>&nbsp;</td> " & _
	"				<td>&nbsp;</td> " & _
	"			</tr> " & _
	"			<tr> " & _
	"				<td> " & _
	"				<p align=""right"">" & cmpMailStr & "</td> " & _
	"				<td> " & _
	"				&nbsp;</td> " & _
	"			</tr> " & _
	"		</table> " & _
	"		</td> " & _
	"		<td rowspan=""2"" colspan=""2""> " & _
	"		<p align=""center""> " & _
	"		<img name=""rechazocuenta_r7_c5"" src=""" & imgPath & "rechazocuenta_r7_c5.jpg"" width=""17"" height=""92"" border=""0"" alt=""""></td> " & _
	"		<td> " & _
	"		<p align=""center""> " & _
	"		<img src=""" & imgPath & "spacer.gif"" width=""1"" height=""75"" border=""0"" alt=""""></td> " & _
	"	</tr> " & _
	"	<tr> " & _
	"		<td colspan=""3""> " & _
	"		<p align=""center""> " & _
	"		<img name=""rechazocuenta_r8_c2"" src=""" & imgPath & "rechazocuenta_r8_c2.jpg"" width=""508"" height=""17"" border=""0"" alt=""""></td> " & _
	"		<td> " & _
	"		<p align=""center""> " & _
	"		<img src=""" & imgPath & "spacer.gif"" width=""1"" height=""17"" border=""0"" alt=""""></td> " & _
	"	</tr> " & _
	"</table> " & _
	"</body> " & _
	"</html> "
	
	getActivationConfRejBody = msgStr
End Function %>
<!--#include file="agentBottom.asp"-->
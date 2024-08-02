<!--#include file="top.asp" -->
<!--#include file="lang/adminClientsAccess.asp" -->
<!--#include file="adminTradSubmit.asp"-->
<% If Request("CardCode") <> "" Then %>
<script type="text/javascript" src="scr/calendar.js"></script>
<script type="text/javascript" src="scr/lang/calendar-<%=Left(Session("myLng"), 2)%>.js"></script>
<script type="text/javascript" src="scr/calendar-setup.js"></script>
<% End If %>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<script language="javascript">
var uField
function Start(o, page, w, h, s) {
uField = o
OpenWin = this.open(page, "DatePicker", "toolbar=no,menubar=no,location=no,scrollbars="+s+",resizable=yes, width="+w+",height="+h);
}
function setTimeStamp(vardate) {
uField.value = vardate
}
</script>
<% 

If Request("submitCmd") = "uAccess" Then
	set oLic = server.CreateObject("TM.LicenceConnect.LicenceConnection")
	oLic.LicenceServer = licip
	oLic.LicencePort = licport
	
	If Request("AccessFrom") <> "" Then dateFrom = "Convert(datetime,'" & SaveSqlDate(Request("AccessFrom")) & "',120)" Else dateFrom = "NULL"
	If Request("AccessTo") <> "" Then dateTo = "Convert(datetime,'" & SaveSqlDate(Request("AccessTo")) & "',120)" Else dateTo = "NULL"
	If Request("Active") = "Y" Then active = "A" Else active = "N"
	If Request("chkApplyPListAgent") = "Y" Then chkApplyPListAgent = "Y" Else chkApplyPListAgent = "N"
	If Request("Password") <> "" Then
		passAdd1 = ", Password = @password"
		passAdd2 = ", Password"
		passAdd3 = ", @Password"
	End If
	If Request("CatalogFilter") <> "" Then CatalogFilter = "N'" & saveHTMLDecode(Request("CatalogFilter"), False) & "'" Else CatalogFilter = "NULL"
	If Request("CatalogFilterAgent") = "Y" Then CatalogFilterAgent = "Y" Else CatalogFilterAgent = "N"
	If Request("AlterLanID") <> "" Then AlterLanID = "'" & Request("AlterLanID") & "'" Else AlterLanID = "NULL"

	sql = "declare @CardCode nvarchar(15) set @CardCode = N'" & saveHTMLDecode(Request("CardCode"), False) & "' " & _
		"declare @dateFrom datetime set @dateFrom = " & dateFrom & " " & _
		"declare @dateTo datetime set @dateTo = " & dateTo & " " & _
		"declare @QryGroup char(1) set @QryGroup = '" & Active & "' " & _
		"declare @password nvarchar(100) set @password = N'" & oLic.GetEncPwd(saveHTMLDecode(Request("password"), False)) & "' " & _
		"declare @PriceList int set @PriceList = " & Request("cmbPList") & " " & _
		"declare @PListApplyAgent char(1) set @PListApplyAgent = '" & chkApplyPListAgent & "' " & _
		"declare @AlterLanID nvarchar(5) set @AlterLanID = " & AlterLanID & " " & _
		"if exists(select 'A' FROM OLKClientsAccess where CardCode = @CardCode) Begin " & _
		"update olkClientsAccess set AlterLanID = @AlterLanID, PriceList = @PriceList, PListApplyAgent = @PListApplyAgent, accessFrom = @dateFrom, accessTo = @dateTo, " & _
		"CatalogFilter = " & CatalogFilter & ", CatalogFilterAgent = '" & CatalogFilterAgent & "', Status = '" & active & "' " & passAdd1 & " where cardcode = @CardCode " & _
		"End Else Begin " & _
		"insert olkClientsAccess(CardCode" & passAdd2 & ", AlterLanID, PriceList, PListApplyAgent, AccessFrom, AccessTo, CatalogFilter, CatalogFilterAgent, Status) " & _
		"values(@CardCode" & passAdd3 & ", @AlterLanID, @PriceList, @PListApplyAgent, @dateFrom, @dateTo, " & CatalogFilter & ", '" & CatalogFilterAgent& "', '" & active & "') " & _
		"End "
		If Request("ipFrom0") <> "" Then
			varIpFrom = Request("ipFrom0") & "."
			If Request("ipFrom1") <> "" Then varIpFrom = varIpFrom & Request("ipFrom1") & "." Else varIpFrom = varIpFrom & "0."
			If Request("ipFrom2") <> "" Then varIpFrom = varIpFrom & Request("ipFrom2") & "." Else varIpFrom = varIpFrom & "0."
			If Request("ipFrom3") <> "" Then varIpFrom = varIpFrom & Request("ipFrom3") Else varIpFrom = varIpFrom & "0"
			If Request("ipTo0") <> "" Then varIpTo = Request("ipTo0") & "." Else varIpTo = "255."
			If Request("ipTo1") <> "" Then varIpTo = varIpTo & Request("ipTo1") & "." Else varIpTo = varIpTo & "255."
			If Request("ipTo2") <> "" Then varIpTo = varIpTo & Request("ipTo2") & "." Else varIpTo = varIpTo & "255."
			If Request("ipTo3") <> "" Then varIpTo = varIpTo & Request("ipTo3") Else varIpTo = varIpTo & "255"
			sql= sql & " declare @IPIndex int set @IPIndex = IsNULL((select max(IPIndex)+1 from OLKClientsAccessIPS where CardCode = N'" & saveHTMLDecode(Request("CardCode"), False) & "'),0) " & _
			"insert OLKClientsAccessIPS(CardCode, IPIndex, IPFrom, IPTo) " & _
			"values(N'" & saveHTMLDecode(Request("CardCode"), False) & "', @IPIndex, '" & varIpFrom & "', '" & varIpTo & "')"
		End If
	conn.execute(sql)
	
	sql = "delete OLKClientsRGAccess where CardCode = N'" & saveHTMLDecode(Request("CardCode"), False) & "' "
	If Request("RGAccess") <> "" Then
		sql = sql & "insert OLKClientsRGAccess(CardCode, rgIndex) select N'" & saveHTMLDecode(Request("CardCode"), False) & "', Value from OLKCommon.dbo.OLKSplit('" & Request("RGAccess") & "', ', ') "
	End If
	conn.execute(sql)
	
	If Request("btnSave") <> "" Then Response.Redirect "adminClientsAccess.asp"
ElseIf Request("submitCmd") = "delIP" Then
	sql = "delete OLKClientsAccessIPS where CardCode = N'" & saveHTMLDecode(Request("CardCode"), False) & "' and ipIndex = " & Request("ipIndex")
conn.execute(sql)
End If
%>
<style type="text/css">
.style1 {
	text-align: center;
	color: #3F7B96;
	background-color: #F5FBFE;
}
.style2 {
	font-weight: bold;
	background-color: #E1F3FD;
	text-align: center;
}
.style3 {
	background-color: #E1F3FD;
}
.style4 {
	text-align: center;
}
.style5 {
	background-color: #F5FBFE;
}
.style6 {
	background-color: #E2F3FC;
}
.style7 {
	background-color: #E2F3FC;
	text-align: center;
}
.style8 {
	color: #31659C;
}
.style9 {
	font-family: Verdana;
	font-size: xx-small;
	color: #31659C;
}
.style10 {
	background-color: #E2F3FC;
	font-family: Verdana;
	font-size: xx-small;
	color: #31659C;
}
.style11 {
	font-family: Verdana;
}
.style13 {
	color: #4783C5;
}
</style>
</head>

<table border="0" cellpadding="0" width="100%" id="table3">
	<% If Request("CardCode") = "" Then %>
	<tr>
		<td bgcolor="#E1F3FD"><b><font color="#FFFFFF">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><%=getadminClientsAccessLngStr("LttlClientsAccess")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1" color="#4783C5"> 
		<%=getadminClientsAccessLngStr("LttlClientsAccessNote")%></font></td>
	</tr>
	<form method="POST" action="adminClientsAccess.asp">
	<tr>
		<td>
		<table border="0" cellpadding="0" id="table4">
			<tr>
				<td bgcolor="#E2F3FC" style="width: 100px">
				<p><font face="Verdana" size="1" color="#4783C5"><strong>
				<span class="style8"><%=getadminClientsAccessLngStr("DtxtSearch")%></span>&nbsp;</strong></font></td>
				<td colspan="3" bgcolor="#F5FBFE"><input type="text" name="searchStr" size="72" class="input" style="width: 100%;" value="<%=Request("searchStr")%>"></td>
			</tr>
			<tr>
				<td bgcolor="#E2F3FC" style="width: 100px" class="style9">
				<p class="style8">
				<font face="Verdana" size="1"><strong><%=getadminClientsAccessLngStr("DtxtGroup")%>&nbsp; 
				</strong> </font></td>
				<td bgcolor="#F5FBFE">
				<font face="Verdana" size="1" color="#4783C5"> 
				<select size="1" name="GroupCode" class="input" style="width: 200px"><option></option><%
				set cmd = Server.CreateObject("ADODB.Command")
				cmd.ActiveConnection = connCommon
				cmd.CommandType = &H0004
				cmd.CommandText = "DBOLKGetCrdGroups" & Session("ID")
				cmd.Parameters.Refresh()
				cmd("@LanID") = Session("LanID")
				cmd("@CardType") = "C"
				set rs = cmd.execute()
				do while not rs.eof 
				%><option <% If CStr(rs("GroupCode")) = CStr(Request("GroupCode")) Then %>selected<% End If %> value="<%=rs("GroupCode")%>"><%=myHTMLEncode(rs("GroupName"))%></option><% rs.movenext
				loop %></select></font></td>
				<td bgcolor="#E2F3FC" style="width: 100px" class="style9">
				<font face="Verdana" size="1"><strong><%=getadminClientsAccessLngStr("DtxtType")%>&nbsp;</strong></font></td>
				<td width="100" bgcolor="#F5FBFE">
				<font face="Verdana" size="1" color="#4783C5"> 
				<select size="1" name="CardType" class="input" style="width: 100px">
				<option value=""><%=getadminClientsAccessLngStr("DtxtAll")%></option>
				<option <% If Request("CardType") = "C" or Request.Form.Count = 0 Then %>selected<% End If %> value="C"><%=getadminClientsAccessLngStr("DtxtClients")%></option>
				<option <% If Request("CardType") = "L" Then %>selected<% End If %> value="L"><%=getadminClientsAccessLngStr("DtxtLeads")%></option>
				</select></font></td>
			</tr>
			<tr>
				<td bgcolor="#E2F3FC" style="width: 100px" class="style9">
				<font face="Verdana" size="1"><strong><%=getadminClientsAccessLngStr("DtxtCountry")%> 
				</strong> 
				</font></td>
				<td bgcolor="#F5FBFE">
				<font face="Verdana" size="1" color="#4783C5"> 
				<select size="1" name="Country" class="input" style="width: 200px"><option></option><%
				sql = "select Code, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRY', 'Name', Code, Name) Name from ocry where exists(select 'A' from ocrd where Country = ocry.Code or MailCountr = ocry.Code) order by name asc"
				set rs = conn.execute(sql)
				do while not rs.eof 
				%><option <% If rs("Code") = Request("Country") Then %>selected<% End If %> value="<%=rs("Code")%>"><%=myHTMLEncode(rs("Name"))%></option><% rs.movenext
				loop %></select></font></td>
				<td bgcolor="#E2F3FC" style="width: 100px" class="style9">
				<font face="Verdana" size="1"><strong><%=getadminClientsAccessLngStr("DtxtState2")%>&nbsp;
				</strong>
				</font></td>
				<td width="100" bgcolor="#F5FBFE">
				<font face="Verdana" size="1" color="#4783C5"> 
				<select size="1" name="cmbStatus" class="input" style="width: 100px">
				<option></option>
				<option <% If Request("cmbStatus") = "A" Then %>selected<% End If %> value="A"><%=getadminClientsAccessLngStr("LtxtActives")%></option>
				<option <% If Request("cmbStatus") = "N" Then %>selected<% End If %> value="N"><%=getadminClientsAccessLngStr("LtxtNotActives")%></option>
				</select></font></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>
		<table cellpadding="0" cellspacing="2" border="0" width="100%">
			<tr>
				<td width="75"><input type="submit" value="<%=getadminClientsAccessLngStr("DtxtSearch")%>" name="B3" class="OlkBtn"></td>
				<td><hr size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
	<input type="hidden" name="search" value="Y">
	</form>
	<% End If %>
	<% If Request("search") = "Y" Then
	
	If Request("GroupCode") <> "" Then GroupCode = " and GroupCode = " & Request("GroupCode")
	If Request("Country") <> "" Then Country = " and Country = '" & Request("Country") & "'"
	If Request("cmbStatus") <> "" Then
		Active = " and IsNull(T1.Status, 'N') = '" & Request("cmbStatus") & "'"
	End If

	If Request("CardType") = "" Then
		TypeStr = "<> 'S'"
	Else
		TypeStr = "= '" & Request("CardType") & "'"
	End If
	sql = "select T0.CardCode, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRD', 'CardName', T0.CardCode, T0.CardName) CardName, T1.Status " & _
	"from ocrd T0 " & _
	"left outer join OLKClientsAccess T1 on T1.CardCode = T0.CardCode " & _
	"where T0.CardType " &  TypeStr & " and (T0.CardCode like N'%" & saveHTMLDecode(Request("searchStr"), False) & _
	"%' or OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRD', 'CardName', T0.CardCode, T0.CardName) collate database_default like N'%" & saveHTMLDecode(Request("searchStr"), False) & "%')" & GroupCode & Country & Active
	rs.close
	rs.PageSize = 10
	rs.CacheSize = 10
	rs.open sql, conn, 3, 1
	If Request("page") = "" Then
		iPageCurrent = 1
	Else
		iPageCurrent = Request("page")
	End If
	
	nPageCount = rs.PageCount
	iNextCount = iPageCurrent
	iCurMax = nPageCount/15
	iCurNext = 0
	
	do while iNextCount > 0
		iNextCount = iNextCount - 15
		iCurNext = iCurNext + 1
	loop
	If iCurMax - CInt(iCurMax) > 0 Then iCurMax = CInt(iCurMax) + 1

	fromI = (iCurNext*15)-14
	toI = (iCurNext*15)
	
	If iCurMax <= iCurNext Then toI = nPageCount
	If nPage > nPageCount Then nPage = nPageCount
	If nPage < 1 Then nPage = 1
	
	If Not rs.eof then rs.AbsolutePage = iPageCurrent %>
	<tr>
		<td bgcolor="#E1F3FD"><b>&nbsp;<font face="Verdana" size="1" color="#31659C"><%=getadminClientsAccessLngStr("DtxtSearch")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1" color="#4783C5"> 
		<%=getadminClientsAccessLngStr("LttlSearchNote")%></font></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="600" id="table6">
			<tr>
				<td width="13" class="style3">&nbsp;</td>
				<td class="style2"><font size="1" face="Verdana" color="#31659C"><%=getadminClientsAccessLngStr("DtxtCode")%></font></td>
				<td class="style2"><font size="1" face="Verdana" color="#31659C"><%=getadminClientsAccessLngStr("DtxtName")%></font></td>
				<td class="style2"><font face="Verdana" size="1" color="#31659C"><%=getadminClientsAccessLngStr("DtxtActive")%></font></td>
			</tr>
			<% If Not rs.eof then
			for i = 0 to rs.PageSize %>
			<tr bgcolor="#F3FBFE">
				<td width="13">
				<p align="justify">
				<font face="Verdana" size="1" color="#4783C5">
				<a href="javascript:goBP('<%=Replace(rs("CardCode"), "'", "\'")%>');">
				<img border="0" src="images/<%=Session("rtl")%>flechaselec.gif" width="15" height="13"></a></font></td>
				<td>
				<p align="justify"><font face="Verdana" size="1" color="#4783C5"><%=rs("CardCode")%></font></td>
				<td>
				<p align="justify"><font face="Verdana" size="1" color="#4783C5"><%=rs("CardName")%></font></td>
				<td class="style4"><font face="Verdana" size="1" color="#4783C5">
				<% If rs("Status") = "A" Then %><%=getadminClientsAccessLngStr("DtxtYes")%><% Else %><%=getadminClientsAccessLngStr("DtxtNo")%><% End If %></font></td>
			</tr>
			<% rs.movenext
			If Rs.Eof Then Exit For
			next %>
			<tr>
				<td colspan="4">
				<table cellpadding="0" border="0" width="100%">
				<tr>
					<td>
					<table border="0" cellpadding="0" width="100%" id="table6">
						<% if rs.PageCount > 1 Then %>
						<tr bgcolor="#F5FBFE">
							<td colspan="4" dir="ltr" align="<% If Session("rtl") = "" Then %>left<% Else %>right<% End If %>">
							<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table7">
								<tr>
									<td width="15">
									<% If iCurNext > 1 Then %><a href="javascript:goAccessPage(<%= ((iCurNext-1)*15) %>);"><img border="0" src="images/prevAll.gif" width="12" height="13"></a><% End If %></td>
									<td width="15">
									<% If iPageCurrent > 1 Then %><a href="javascript:goAccessPage(<%=iPageCurrent-1%>)"><img border="0" src="images/flechaselec2.gif" width="15" height="13"></a><% End If %></td>
									<td><font face="Verdana" size="1">
									<p align="center"><b>&nbsp;<% For i = fromI To toI %>
									<% If i <> CInt(iPageCurrent) Then %><a href="javascript:goAccessPage(<%=i%>)"><font color="#4783C5"><% Else %><font color="red"><% End If %><%=i%><% If i <> CInt(iPageCurrent) Then %></font></a><font color="#4783C5"><% End If %>&nbsp;
									<% Next %></font></b></td>
									<td width="15">
									<% If CInt(iPageCurrent) < rs.PageCount Then %><a href="javascript:goAccessPage(<%=iPageCurrent+1%>)"><img border="0" src="images/flechaselec.gif" width="15" height="13"></a><% End If %></td>
									<td width="15">
									<% If iCurNext < iCurMax Then %><a href="javascript:goAccessPage(<%= (iCurNext*15)+1 %>);"><img border="0" src="images/nextAll.gif" width="12" height="13"></a><% End If %></td>
								</tr>
							</table>
							</td>
						</tr>
						<% 
						End If
						Else %>
						<tr bgcolor="#F5FBFE"><td colspan="4" width="100%" align="center"><font face="Verdana" size="1" color="#4783C5"><%=getadminClientsAccessLngStr("LtxtNoData")%></font></td>
						</tr>
						<% End If %>
						</table>
					</td>
				</tr>
				</table>
				</td>
			</tr>
		</table>
		</td>
	</tr>
	<form name="frmGoCAccess" method="post" action="adminClientsAccess.asp">
	<% For each itm in Request.Form
		If itm <> "page" Then %>
	<input type="hidden" name="<%=itm%>" value="<%=Request(itm)%>">
	<% End If
	Next %>
	<input type="hidden" name="page" value="">
	</form>
	<script language="javascript">
	function goAccessPage(p) { document.frmGoCAccess.page.value = p; document.frmGoCAccess.submit(); }
	</script>
	<% End If 
	If Request("CardCode") <> "" Then 
	sql = "select T0.CardCode, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRD', 'CardName', T0.CardCode, T0.CardName) CardName, T1.AlterLanID, T1.Status, T1.PriceList, T1.PListApplyAgent, AccessFrom, AccessTo, " & _
	"Case When Exists(select 'A' from olkClientsAccess where CardCode = T0.CardCode) Then 'Y' Else 'N' End Verfy, " & _
	"IsNull(T1.CatalogFilter, '') CatalogFilter, T1.CatalogFilterAgent, T0.DocEntry " & _
	"from OCRD T0 " & _
	"left outer join olkClientsAccess T1 on T1.CardCode = T0.CardCode where T0.CardCode = N'" & saveHTMLDecode(Request("CardCode"), False) & "'"
	set rs = conn.execute(sql) 
	If rs("Verfy") = "N" Then
		newPwd = True
	ElseIf rs("Verfy") = "Y" Then
		newPwd = False
	End If
	CatalogFilter = rs("CatalogFilter")
	CatalogFilterAgent = rs("CatalogFilterAgent")
	DocEntry = rs("DocEntry") %>
	<form method="POST" action="adminClientsAccess.asp" name="frmEdit" onsubmit="return valEditFrm()">
	<tr>
		<td bgcolor="#E1F3FD"><b><font color="#FFFFFF">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><%=getadminClientsAccessLngStr("LttlAccessProp")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1" color="#4783C5"> 
		<%=getadminClientsAccessLngStr("LtxtAccessPropNote")%></font></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td width="800">
				<table border="0" width="100%" id="table12" cellpadding="0">
					<tr>
						<td class="style10"><font size="1" face="Verdana" color="#31659C">
						<strong><%=getadminClientsAccessLngStr("DtxtCode")%></strong></font></td>
						<td class="style5"><font size="1" face="Verdana" color="#4783C5"><%=rs("CardCode")%></font>&nbsp;</td>
					</tr>
					<tr>
						<td class="style10"><font face="Verdana" size="1" color="#31659C">
						<strong><%=getadminClientsAccessLngStr("DtxtName")%>&nbsp;</strong></font></td>
						<td class="style5"><font face="Verdana" size="1" color="#4783C5"><%=rs("CardName")%>&nbsp;</font>&nbsp;</td>
					</tr>
					<tr>
						<td class="style10"><font face="Verdana" size="1" color="#31659C">
						<strong><%=getadminClientsAccessLngStr("DtxtPList")%></strong></font></td>
						<td class="style5"><select size="1" name="cmbPList" class="input" style="width: 180px;">
						<option value="-9"><%=getadminClientsAccessLngStr("LtxtClientDef")%></option>
						<% 
					    set cmd = Server.CreateObject("ADODB.Command")
						cmd.ActiveConnection = connCommon
						cmd.CommandType = &H0004
						cmd.CommandText = "DBOLKGetPriceList" & Session("ID")
						cmd.Parameters.Refresh()
						cmd("@LanID") = Session("LanID")
						set rd = cmd.execute()
						do while not rd.eof %>
						<option <% If rs("PriceList") = rd(0) Then %>selected<% End If %> value="<%=rd(0)%>"><%=myHTMLEncode(rd(1))%></option>
						<% rd.movenext
						loop %>
						</select></td>
					</tr>
					<tr>
						<td class="style10">&nbsp;</td>
						<td class="style5"><font face="Verdana" size="1" color="#31659C">
						<span class="style11"><font size="1"><span class="style13">
						<input type="checkbox" class="noborder" <% If rs("PListApplyAgent") = "Y" Then %>checked<% End If %> name="chkApplyPListAgent" value="Y" id="chkApplyPListAgent"></span></font></span></font><font face="Verdana" size="1"><label for="chkApplyPListAgent"><span class="style13"><%=getadminClientsAccessLngStr("LtxtApplyPListAgent")%></span></label></font></td>
					</tr>
					<tr>
						<td class="style10"><font face="Verdana" size="1" color="#31659C">
						<strong><%=getadminClientsAccessLngStr("LtxtLng")%></strong></font></td>
						<td class="style5">
							<select name="AlterLanID" class="input" style="width: 180px;">
							<option value=""><%=getadminClientsAccessLngStr("LtxtNatLng")%></option>
							<% For i = 0 to UBound(myLanIndex) %>
							<option <% If rs("AlterLanID") = myLanIndex(i)(0) Then %>selected<% End If %> value="<%=myLanIndex(i)(0)%>"><%=myLanIndex(i)(1)%></option>
							<% Next %>
							</select></td>
					</tr>
					<tr>
						<td class="style10">&nbsp;</td>
						<td class="style5">
							<span class="style11">
						<font size="1"><span class="style13"><input type="checkbox" <% If rs("Status") = "A" Then %>checked<% End If %> name="Active" value="Y" class="noborder" id="chkActive" style="height: 20px"></span></font></span><font face="Verdana" size="1"><label for="chkActive"><span class="style13"><%=getadminClientsAccessLngStr("DtxtActive")%></span></label></font></td>
					</tr>
					<tr>
						<td class="style10"><font face="Verdana" size="1" color="#31659C">
								<strong><%=getadminClientsAccessLngStr("LtxtNewPwd")%></strong></font></td>
						<td class="style5">
						<input type="password" name="password" size="20" class="input" onkeydown="return chkMax(event, this, 20);"></td>
					</tr>
					<tr>
						<td class="style10"><font face="Verdana" size="1" color="#31659C">
								<strong><%=getadminClientsAccessLngStr("LtxtPwdConf")%></strong></font></td>
						<td class="style5">
						<input type="password" name="cpassword" size="20" class="input" onkeydown="return chkMax(event, this, 20);"></td>
					</tr>
					<tr>
						<td class="style10"><font face="Verdana" size="1" color="#31659C">
						<strong><%=getadminClientsAccessLngStr("LtxtAccessFrom")%></strong></font></td>
						<td class="style5">
							<table cellpadding="0" cellspacing="0" border="0">
								<tr>
									<td><img border="0" src="images/cal.gif" id="btnAccessFrom" width="16" height="16">&nbsp;</td>
									<td><input readonly type="text" name="AccessFrom" id="AccessFrom" size="16" class="input" onclick="btnAccessFrom.click()" value="<%=FormatDate(rs("AccessFrom"), False)%>"></td>
									<td><img border="0" src="images/remove.gif" width="16" height="16" style="cursor: hand" onclick="document.frmEdit.AccessFrom.value=''"></td>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td class="style10"><font face="Verdana" size="1" color="#31659C"><strong><%=getadminClientsAccessLngStr("LtxtAccessTo")%></strong></font></td>
						<td class="style5">
							<table cellpadding="0" cellspacing="0" border="0">
								<tr>
									<td><img border="0" src="images/cal.gif" id="btnAccessTo" width="16" height="16">&nbsp;</td>
									<td><input readonly type="text" name="AccessTo" id="AccessTo" size="16" class="input" onclick="btnAccessTo.click()" value="<%=FormatDate(rs("AccessTo"), False)%>"></td>
									<td><img border="0" src="images/remove.gif" width="16" height="16" onclick="document.frmEdit.AccessTo.value=''"></td>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td class="style10" valign="top" style="padding-top: 2px;"><font face="Verdana" size="1" color="#31659C"><strong><%=getadminClientsAccessLngStr("LtxtIPFilter")%></strong></font></td>
						<td class="style5">
							<table border="0" cellpadding="0" width="400">
								<tr>
									<td class="style7"><font face="Verdana" size="1" color="#31659C">
									<strong><%=getadminClientsAccessLngStr("LtxtIPFrom")%></strong></font></td>
									<td style="width: 16px" class="style6">&nbsp;</td>
									<td class="style7"><font face="Verdana" size="1" color="#31659C">
									<strong><%=getadminClientsAccessLngStr("LtxtIPTo")%></strong></font></td>
									<td class="style6" style="width: 16px">&nbsp;</td>
								</tr>
								<% sql = "select * from OLKClientsAccessIPS where cardcode = N'" & saveHTMLDecode(Request("CardCode"), False) & "'"
								set rs = conn.execute(sql)
								do while not rs.eof %>
								<tr>
									<td class="style5"><font face="Verdana" size="1" color="#4783C5"><%=rs("IPFrom")%>&nbsp;</font></td>
									<td style="width: 16px" class="style5">&nbsp;</td>
									<td class="style5"><font face="Verdana" size="1" color="#4783C5"><%=rs("IPTo")%>&nbsp;</font></td>
									<td class="style5" style="width: 16px"><a href="javascript:if(confirm('<%=getadminClientsAccessLngStr("LtxtConfDelIP")%>'.replace('{0}', '<%=rs("IPFrom")%>').replace('{1}', '<%=rs("IPTo")%>')))doMyLink('adminClientsAccess.asp', 'CardCode=<%=Request("CardCode")%>&ipIndex=<%=rs("IPIndex")%>&submitCmd=delIP&searchStr=<%=Request("searchStr")%>&GroupCode<%=Request("GroupCode")%>&Country=<%=Request("Country")%>', '');"><img border="0" src="images/remove.gif" width="16" height="16"></a></td>
								</tr>
								<% rs.movenext
								loop %>
								<tr>
									<td class="style5">
									<font size="3">
									<input name="IpFrom0" id="IpFrom0" size="3" class="input" onfocus="this.select()" style="font-weight: 700" onkeyup="chkNum(this)">.<input name="IpFrom1" id="IpFrom1" size="3" class="input" onfocus="this.select()" style="font-weight: 700" onkeyup="chkNum(this)">.<input name="IpFrom2" id="IpFrom2" size="3" class="input" onfocus="this.select()" style="font-weight: 700" onkeyup="chkNum(this)">.<input name="IpFrom3" id="IpFrom3" size="3" class="input" onfocus="this.select()" style="font-weight: 700" onkeyup="chkNum(this)"></font></td>
									<td style="width: 16px" class="style1">
									<strong>-</strong></td>
									<td class="style5">
									<font size="3">
									<input name="IpTo0" id="IpTo0" size="3" class="input" onfocus="this.select()" style="font-weight: 700" onkeyup="chkNum(this)">.<input name="IpTo1" id="IpTo1" size="3" class="input" onfocus="this.select()" style="font-weight: 700" onkeyup="chkNum(this)">.<input name="IpTo2" id="IpTo2" size="3" class="input" onfocus="this.select()" style="font-weight: 700" onkeyup="chkNum(this)">.<input name="IpTo3" id="IpTo3" size="3" class="input" onfocus="this.select()" style="font-weight: 700" onkeyup="chkNum(this)"></font></td>
									<td class="style5" style="width: 16px">&nbsp;</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td class="style10" valign="top" style="padding-top: 2px;" rowspan="2">
						<font face="Verdana" size="1" color="#31659C">
						<strong><%=getadminClientsAccessLngStr("LtxtCatFlt")%><br>(</strong>ItemCode not in<strong>)</strong></font></td>
						<td class="style5">
							<table cellpadding="0" cellspacing="0" border="0" width="100%">
								<tr>
									<td rowspan="2">
										<textarea dir="ltr" rows="18" name="CatalogFilter" cols="63" class="input" style="width: 100%; " onkeypress="javascript:document.frmEdit.btnVerfyFilter.src='images/btnValidate.gif';document.frmEdit.btnVerfyFilter.style.cursor = 'hand';;document.frmEdit.valCatalogFilter.value='Y';"><%=Server.HTMLEncode(CatalogFilter)%></textarea>
									</td>
									<td valign="top" width="1">
										<img style="cursor: hand; " src="images/qry_note.gif" id="btnNoteFilter" alt="<%=getadminClientsAccessLngStr("DtxtDefinition")%>" onclick="javascript:doFldNote(7, 'CatalogFilter', <%=DocEntry%>, null);">
									</td>
								</tr>
								<tr>
									<td valign="bottom" width="1">
										<img src="images/btnValidateDis.gif" id="btnVerfyFilter" alt="<%=getadminClientsAccessLngStr("DtxtValidate")%>" onclick="javascript:if (document.frmEdit.valCatalogFilter.value == 'Y')VerfyQuery();">
										<input type="hidden" name="valCatalogFilter" value="N">						</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td class="style5">
						<font face="Verdana" size="1" color="#31659C">
								<input type="checkbox" <% If CatalogFilterAgent = "Y" Then %>checked<% End If %> class="noborder" name="CatalogFilterAgent" id="CatalogFilterAgent" value="Y"><label for="CatalogFilterAgent"><%=getadminClientsAccessLngStr("LtxtCatalogFilterAgen")%></label></font></td>
					</tr>
					<tr>
						<td class="style10" valign="top" style="padding-top: 2px;">
						<font face="Verdana" size="1" color="#31659C"><strong><%=getadminClientsAccessLngStr("LtxtRepGrp")%></strong></font></td>
						<td class="style5">
						<ilayer name="scroll1" width=500 height=300 clip="0,0,170,150">
						<layer name="scroll2" width=500 height=300 bgColor="white">
						<div id="scroll3" style="width:500;height:300px;overflow:auto">
						<table cellpadding="0" cellspacing="0" border="0">
						<% sql = "select T0.rgIndex, T0.rgName, Case When T1.rgIndex is not null then 'Y' Else 'N' End Verfy, SuperUser " & _
						"from OLKRG T0 " & _
						"left outer join OLKClientsRGAccess T1 on T1.CardCode = N'" & saveHTMLDecode(Request("CardCode"), False) & "' and T1.rgIndex = T0.rgIndex " & _
						"where T0.UserType = 'C' " & _
						"order by 2 asc"
						set rs = conn.execute(sql)
						do while not rs.eof %>
						<tr>
							<td><input type="checkbox" class="noborder"<% If rs("SuperUser") = "Y" Then %> <% If rs("Verfy") = "Y" Then %> checked <% End If %> name="RGAccess" id="RGAccess<%=rs("rgIndex")%>"<% Else %>checked disabled<% End If %> value="<%=rs("rgIndex")%>"><% If rs("SuperUser") = "N" Then %><input type="hidden" name="RGAccess" value="<%=rs("rgIndex")%>"><% End If %></td>
							<td><font face="Verdana" size="1" color="#31659C"><label for="RGAccess<%=rs("rgIndex")%>"><%=rs("rgName")%></label></font></td>
						</tr>
						<% rs.movenext
						loop %>
						</table>
						</div>
						</layer>
						</ilayer>
						</td>
					</tr>
					</table>
				</td>
			</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td>
		<table cellpadding="0" cellspacing="2" border="0" width="100%">
		<tr>
			<td width="75"><input type="submit" value="<%=getadminClientsAccessLngStr("DtxtApply")%>" name="btnApply" class="OlkBtn"></td>
			<td width="75"><input type="submit" value="<%=getadminClientsAccessLngStr("DtxtSave")%>" name="btnSave" class="OlkBtn"></td>
			<td><hr size="1"></td>
			<td width="75"><input type="button" value="<%=getadminClientsAccessLngStr("DtxtCancel")%>" name="btnCancel" class="OlkBtn" onclick="javascript:window.location.href='adminClientsAccess.asp'"></td>
		</tr>
	</table>
	</td>
	</tr>
	<input type="hidden" name="CardCode" value="<%=myHTMLEncode(Request("CardCode"))%>">
	<input type="hidden" name="Country" value="<%=Request("Country")%>">
	<input type="hidden" name="GroupCode" value="<%=Request("GroupCode")%>">
	<input type="hidden" name="searchStr" value="<%=Request("searchStr")%>">
	<input type="hidden" name="submitCmd" value="uAccess">
	</form>
	<script language="javascript">
	function valEditFrm() 
	{
		var Validation = true
		<% If newPwd Then %>
		if (document.frmEdit.password.value == "") 
		{
			alert("<%=getadminClientsAccessLngStr("LtxtValNewUPwd")%>");
			document.frmEdit.password.focus();
			return false 
		}
		<% End If %>
		if (document.frmEdit.password.value != document.frmEdit.cpassword.value) 
		{
			alert("<%=getadminClientsAccessLngStr("LtxtValPwdConf")%>");
			document.frmEdit.password.value = "";
			document.frmEdit.cpassword.value = "";
			document.frmEdit.password.focus();
			return false 
		}
		if (document.frmEdit.CatalogFilter.value != '' && document.frmEdit.valCatalogFilter.value == 'Y')
		{
			alert("<%=getadminClientsAccessLngStr("LtxtValQryVal")%>");
			document.frmEdit.btnVerfyFilter.focus();
			return false; 
		}
	}
	
	function chkNum(fld)
	{
		if (fld.value == '.') fld.value = '';
		if (fld.value != '')
		{
			var varInStr = fld.value.indexOf('.');
			if (fld.value.length == 3 || varInStr != -1)
			{
				fld.value = fld.value.replace('.', '');
				var varFld = Left(fld.name, fld.name.length-1);
				var varNextFld = parseInt(Right(fld.name, 1))+1;
				
				switch (varFld)
				{
					case 'IpFrom':
						switch (varNextFld)
						{
							case 1:
								document.frmEdit.IpFrom1.focus();
								break;
							case 2:
								document.frmEdit.IpFrom2.focus();
								break;
							case 3:
								document.frmEdit.IpFrom3.focus();
								break;
						}
						break;
					case 'IpTo':
						switch (varNextFld)
						{
							case 1:
								document.frmEdit.IpTo1.focus();
								break;
							case 2:
								document.frmEdit.IpTo2.focus();
								break;
							case 3:
								document.frmEdit.IpTo3.focus();
								break;
						}
						break;
				}
			}
			if (!IsNumeric(fld.value) && fld.value != '' && varInStr == -1)
			{
				alert('<%=getadminClientsAccessLngStr("DtxtValNumVal")%>');
				fld.value = '';
				fld.focus();
			}
			else if (parseInt(fld.value) > 255 || parseInt(fld.value) < 0)
			{
				alert('<%=getadminClientsAccessLngStr("LtxtValFldLimit")%>');
				fld.value = '';
				fld.focus();
			}
		}
	}
	</script>
	<% End If %>
	</table>
<script language="javascript">
<!--
function VerfyQuery()
{
	document.frmVerfyQuery.Query.value = document.frmEdit.CatalogFilter.value;
	if (document.frmVerfyQuery.Query.value != '')
	{
		document.frmVerfyQuery.submit();
	}
	else
	{
		VerfyQueryVerified();
	}
}
function VerfyQueryVerified()
{
	//document.frmEdit.btnVerfyQuery.disabled = true;
	document.frmEdit.btnVerfyFilter.src='images/btnValidateDis.gif'
	document.frmEdit.btnVerfyFilter.style.cursor = '';
	document.frmEdit.valCatalogFilter.value='N';
}
//-->
</script>
<form name="frmVerfyQuery" action="verfyQuery.asp" method="post" target="iVerfyQuery">
	<iframe id="iVerfyQuery" name="iVerfyQuery" style="display: none" src=""></iframe>
	<input type="hidden" name="type" value="ClientCatFilter">
	<input type="hidden" name="Query" value="">
	<input type="hidden" name="by" value="">
	<input type="hidden" name="parent" value="Y">
</form>
<% If Request("CardCode") <> "" Then %>
<script type="text/javascript">
    Calendar.setup({
        inputField     :    "AccessFrom",     // id of the input field
        ifFormat       :    "<%=GetCalendarFormatString%>",      // format of the input field
        button         :    "btnAccessFrom",  // trigger for the calendar (button ID)
        align          :    "Bl",           // alignment (defaults to "Bl")
        singleClick    :    true
    });
    Calendar.setup({
        inputField     :    "AccessTo",     // id of the input field
        ifFormat       :    "<%=GetCalendarFormatString%>",      // format of the input field
        button         :    "btnAccessTo",  // trigger for the calendar (button ID)
        align          :    "Br",           // alignment (defaults to "Bl")
        singleClick    :    true
    });
</script>
<% Else %>
<script language="javascript">
function goBP(bp)
{
	document.frmBP.CardCode.value = bp;
	document.frmBP.submit();
}
</script>
<form name="frmBP" action="adminClientsAccess.asp" method="post">
<input type="hidden" name="CardCode" value="">
<input type="hidden" name="searchStr" value="<%=Request("searchStr")%>">
<input type="hidden" name="GroupCode" value="<%=Request("GroupCode")%>">
<input type="hidden" name="Country" value="<%=Request("Country")%>">
</form>
<% End If %><!--#include file="bottom.asp" -->
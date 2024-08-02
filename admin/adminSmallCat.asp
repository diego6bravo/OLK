<!--#include file="top.asp" -->
<!--#include file="lang/adminSmallCat.asp" -->
<!--#include file="adminTradSubmit.asp"-->
<style type="text/css">
.style1 {
	font-weight: bold; background-color: #E1F3FD;
}
.style2 {
	background-color: #E1F3FD;
}
.style3 {
	background-color: #F3FBFE;
}
.style4 {
	font-family: Verdana;
	font-size: xx-small;
	color: #31659C;
}
</style>

<% 
CatType = "R"
If Request("CatType") <> "" Then CatType = Request("CatType")
If Request("editID") = "" and Request("new") <> "Y" Then %>
<form method="POST" action="adminSmallCatSubmit.asp" name="frmSmallCat">
<input type="hidden" name="cmd" value="upd">
<table border="0" cellpadding="0" width="100%">
	<tr>
		<td bgcolor="#E1F3FD"><b><font color="#FFFFFF">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><%=getadminSmallCatLngStr("LttlSmallCat")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1">
		<font color="#4783C5"><%=getadminSmallCatLngStr("LttlSmallCatNote")%></font></font></p>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0">
			<tr>
				<td bgcolor="#E2F3FC" style="width: 200px"><b>
				<font size="1" face="Verdana" color="#31659C">
				<%=getadminSmallCatLngStr("DtxtType")%>&nbsp;</font></b></td>
				<td valign="top" bgcolor="#F3FBFE">
				<select size="1" name="CatType" class="input" onchange="javascript:window.location.href='adminSmallCat.asp?CatType='+this.value;">
				<option value="R"><%=getadminSmallCatLngStr("DtxtRegular")%></option>
				<option value="I" <% If CatType = "I" Then %>selected<% End If %>><%=getadminSmallCatLngStr("DtxtItem")%></option>
				</select></td>
			</tr>
		</table>
		<table border="0" cellpadding="0" width="100%">
			<tr style="text-align: center; color: #31659C; font-size: x-small; font-family: Verdana;font-weight: bold; background-color: #E1F3FD;">
				<td style="width: 16px;">&nbsp;</td>
				<td><%=getadminSmallCatLngStr("DtxtName")%></td>
				<td><%=getadminSmallCatLngStr("LtxtSubtitle")%></td>
				<td><%=getadminSmallCatLngStr("LtxtDirection")%></td>
				<td><%=getadminSmallCatLngStr("DtxtQuery")%></td>
				<td><%=getadminSmallCatLngStr("DtxtActive")%></td>
				<td style="width: 16px;">&nbsp;</td>
			</tr>
			<% 
			set rs = Server.CreateObject("ADODB.RecordSet")
			sql = "select ID, Name, SubTitle, Direction, Query, [Status] from OLKSmallCat where CatType = '" & CatType & "'"
			rs.open sql, conn, 3, 1
			do while not rs.eof %>
			<tr style="background-color: #F3FBFE;text-align : center; color: #31659C; font-size: x-small; font-family: Verdana;">
				<td style="width: 16px;"><a href='adminSmallCat.asp?CatType=<%=CatType%>&editID=<%=rs("ID")%>'><img border="0" src="images/<%=Session("rtl")%>flechaselec.gif"></a></td>
				<td><%=rs("Name")%><input type="hidden" name="ID" value="<%=rs("ID")%>"></td>
				<td><%=rs("SubTitle")%></td>
				<td><% Select Case rs("Direction")
					Case "V" %><%=getadminSmallCatLngStr("LtxtVertical")%><%
					Case "H" %><%=getadminSmallCatLngStr("LtxtHorizontal")%><%
				End Select %></td>
				<td style="text-align: center;"><img src="images/eye_icon.gif" dir="ltr" border="0" title="<%=Server.HTMLEncode(RS("Query"))%>"></td>
				<td style="text-align: center;"><input class="noborder" type="checkbox" <% If rs("Status") = "A" Then %>checked<% End If %> name="chkStatus<%=rs("ID")%>" value="Y"></td>
				<td style="width: 16px;"><a href="javascript:if(confirm('<%=getadminSmallCatLngStr("LtxtConfDelSmallCat")%>'.replace('{0}', '<%=Replace(Server.HTMLEncode(Rs("Name")),"'","\'")%>')))window.location.href='adminSmallCatSubmit.asp?cmd=del&ID=<%=rs("ID")%>&CatType=<%=CatType%>';"><img border="0" src="images/remove.gif" width="16" height="16"></a></td>
			</tr>
			<% rs.movenext
			loop %>
		</table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr>
				<td width="77">
				<input type="submit" value="<%=getadminSmallCatLngStr("DtxtSave")%>" <% If rs.recordcount = 0 then %>disabled<% End If %> name="B1" class="OlkBtn"></td>
				<td width="77">
				<input type="button" value="<%=getadminSmallCatLngStr("DtxtNew")%>" name="btnNew" class="OlkBtn" onclick="javascript:window.location.href='adminSmallCat.asp?CatType=<%=CatType%>&new=Y'"></td>
				<td><hr color="#0D85C6" size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
</table>
</form>
<% Else %>
<script type="text/javascript">
<!--
function valSmallCat()
{
	if (document.frmEditSmallCat.strName.value == '')
	{
		alert('<%=getadminSmallCatLngStr("LtxtValName")%>');
		document.frmEditSmallCat.strName.focus();
		return false;
	}
	
	if (document.frmEditSmallCat.Query.value == '')
	{
		alert('<%=getadminSmallCatLngStr("LtxtValQuery")%>');
		document.frmEditSmallCat.Query.focus();
		return false;
	}
	
	if (document.frmEditSmallCat.valQuery.value == 'Y')
	{
		alert('<%=getadminSmallCatLngStr("LtxtValQueryVal")%>');
		document.frmEditSmallCat.strName.focus();
		return false;
	}
	return true;
}
function VerfyQuery()
{
	if (document.frmEditSmallCat.Query.value != '')
	{
		$.post('verfyQueryFetch.asp', { Type: 'SmallCat', CatType: document.frmEditSmallCat.CatType.value, Query: document.frmEditSmallCat.Query.value }, function(data)
		{
			if (data == 'ok')
			{
				Verified();
			}
			else
			{
				alert(data);
			}
		});
	}
	else
	{
		Verified();
	}
}
function Verified()
{
	document.frmEditSmallCat.btnVerfyQuery.src='images/btnValidateDis.gif'
	document.frmEditSmallCat.btnVerfyQuery.style.cursor = '';
	document.frmEditSmallCat.valQuery.value='N';
}


//-->
</script>
<form method="POST" action="adminSmallCatSubmit.asp" name="frmEditSmallCat" onsubmit="return valSmallCat();">
<input type="hidden" name="cmd" value="edit">
<input type="hidden" name="editID" value="<%=Request("editID")%>">
<input type="hidden" name="CatType" value="<%=CatType%>">
<table border="0" cellpadding="0" width="100%">
	<tr>
		<td bgcolor="#E1F3FD">&nbsp;<b><font size="1" face="Verdana" color="#31659C"><% If Request("editID") = "" Then %><%=getadminSmallCatLngStr("LttlAddSmallCat")%><% Else %><%=getadminSmallCatLngStr("LttlEditSmallCat")%><% End If %></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1">
		<font color="#4783C5"><%=getadminSmallCatLngStr("LttlAddEditSmallCatNo")%> </font></font></td>
	</tr>
	<tr>
		<td>
		<% If Request("editID") <> "" Then
			sql = "select Name, IsNull(SubTitle, '') SubTitle, Direction, [Top], Query, Status from OLKSmallCat where ID = " & Request("editID") 
			set rs = conn.execute(sql)
			strName = rs("Name")
			subTitle = rs("subTitle")
			direction = rs("Direction")
			strTop = rs("Top")
			Query = rs("Query")
			strStatus = rs("Status") = "A"
		Else
			strName = ""
			subTitle = ""
			direction = "V"
			strTop = 5
			Query = ""
			strStatus = True %>
		<input type="hidden" name="nameTrad">
		<input type="hidden" name="subtitleTrad">
		<input type="hidden" name="queryDef">
		<% End If %>
		<table border="0" cellpadding="0" width="100%">
			<tr>
				<td>
				<table border="0" cellpadding="0">
					<tr>
						<td bgcolor="#E2F3FC"><b>
						<font size="1" face="Verdana" color="#31659C">
						<%=getadminSmallCatLngStr("DtxtName")%>&nbsp;</font></b></td>
						<td valign="top" class="style3" style="width: 260px">
						<p align="center">
						<table cellpadding="0" cellspacing="0" border="0" width="100%">
							<tr>
								<td><input name="strName" style="width: 100%; " class="input" value="<%=Server.HTMLEncode(strName)%>" size="25" onkeydown="return chkMax(event, this, 50);">
								</td>
								<td width="16"><a href="javascript:doFldTrad('SmallCat', 'ID', '<%=Request("editID")%>', 'alterName', 'T', <% If Request("new") <> "Y" Then %>null<% Else %>document.frmEditSmallCat.nameTrad<% End If %>);"><img src="images/trad.gif" alt="<%=getadminSmallCatLngStr("DtxtTranslate")%>" border="0"></a></td>
							</tr>
						</table>
						</td>
					</tr>
					<tr>
						<td bgcolor="#E2F3FC"><b>
						<font size="1" face="Verdana" color="#31659C">
						<%=getadminSmallCatLngStr("LtxtSubtitle")%>&nbsp;</font></b></td>
						<td valign="top" class="style3" style="width: 260px">
						<p align="center">
						<table cellpadding="0" cellspacing="0" border="0" width="100%">
							<tr>
								<td><input name="subTitle" style="width: 100%; " class="input" value="<%=Server.HTMLEncode(subTitle)%>" size="25" onkeydown="return chkMax(event, this, 50);">
								</td>
								<td width="16"><a href="javascript:doFldTrad('SmallCat', 'ID', '<%=Request("editID")%>', 'altersubTitle', 'T', <% If Request("new") <> "Y" Then %>null<% Else %>document.frmEditSmallCat.subtitleTrad<% End If %>);"><img src="images/trad.gif" alt="<%=getadminSmallCatLngStr("DtxtTranslate")%>" border="0"></a></td>
							</tr>
						</table>
						</td>
					</tr>
					<tr>
						<td bgcolor="#E2F3FC">
						<font face="Verdana" size="1" color="#31659C"><strong><%=getadminSmallCatLngStr("LtxtDirection")%></strong></font></td>
						<td valign="top" class="style3">
						<select size="1" name="direction">
						<option <% If direction = "H" Then %>selected<% End If %> value="H"><%=getadminSmallCatLngStr("LtxtHorizontal")%></option>
						<option <% If direction = "V" Then %>selected<% End If %> value="V"><%=getadminSmallCatLngStr("LtxtVertical")%></option>
						</select></td>
					</tr>
					<tr>
						<td bgcolor="#E2F3FC"><b>
						<font face="Verdana" size="1" color="#31659C">
						<%=getadminSmallCatLngStr("DtxtTop")%></font></b></td>
						<td valign="top" class="style3">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td valign="top">
									<input type="text" name="Top" id="Top" size="7" style="text-align:right" class="input" onfocus="this.select();" onmouseup="event.preventDefault()" onkeydown="return chkMax(event, this, 6);" value="<%=strTop%>">
								</td>
								<td valign="middle">
								<table cellpadding="0" cellspacing="0" border="0">
									<tr>
										<td><img src="images/img_nud_up.gif" id="btnTopUp"></td>
									</tr>
									<tr>
										<td><img src="images/spacer.gif"></td>
									</tr>
									<tr>
										<td><img src="images/img_nud_down.gif" id="btnTopDown"></td>
									</tr>
								</table></td>
							</tr>
						</table></td>
					</tr>
					<tr>
						<td valign="top" colspan="2">
						<table border="0" width="100%" cellpadding="0">
							<tr>
								<td valign="top" style="width: 119px" bgcolor="#E2F3FC" class="style4">
								<font size="1" face="Verdana"><strong><%=getadminSmallCatLngStr("DtxtQuery")%></strong><br><b>where ItemCode in (...)</b></font></td>
								<td valign="top" bgcolor="#E2F3FC">
								<table cellpadding="0" cellspacing="0" border="0" width="100%">
									<tr>
										<td rowspan="2">
											<textarea dir="ltr" rows="10" style="width: 100%" name="Query" cols="100" class="input" onkeypress="javascript:document.frmEditSmallCat.btnVerfyQuery.src='images/btnValidate.gif';document.frmEditSmallCat.btnVerfyQuery.style.cursor = 'hand';;document.frmEditSmallCat.valQuery.value='Y';"><%=myHTMLEncode(Query)%></textarea>
										</td>
										<td valign="top" width="1">
											<img style="cursor: hand; " src="images/qry_note.gif" id="btnNoteFilter" alt="<%=getadminSmallCatLngStr("DtxtDefinition")%>" onclick="javascript:doFldNote(25, 'Query', '<%=Request("editID")%>', <% If Request("new") <> "Y" Then %>null<% Else %>document.frmEditSmallCat.queryDef<% End If %>);">
										</td>
									</tr>
									<tr>
										<td valign="bottom" width="1">
											<img src="images/btnValidateDis.gif" id="btnVerfyQuery" alt="|D:txtValidate|" onclick="javascript:if (document.frmEditSmallCat.valQuery.value == 'Y')VerfyQuery();">
											<input type="hidden" name="valQuery" value="N">
										</td>
									</tr>
								</table>
								</td>
							</tr>
							<tr>
								<td valign="top" colspan="2">
								<table cellpadding="0" style="width: 100%">
									<tr>
										<td valign="top" style="width: 119px" bgcolor="#E2F3FC" class="style4">
										<font size="1" face="Verdana">
										<strong><%=getadminSmallCatLngStr("DtxtVariables")%></strong></font></td>
										<td class="style3"><font face="Verdana" size="1" color="#4783C5"><span dir="ltr">@CardCode</span> = <%=getadminSmallCatLngStr("DtxtClientCode")%><br>
										<% If Request("CartType") = "I" Then %><span dir="ltr">@ItemCode</span> = <%=getadminSmallCatLngStr("DtxtItemCode")%><br><% End If %>
										<span dir="ltr">@dbName</span> = <%=getadminSmallCatLngStr("DtxtDB")%><br>
										<span dir="ltr">@LanID</span> = <%=getadminSmallCatLngStr("DtxtLanID")%></font></td>
									</tr>
									<tr>
										<td valign="top" style="width: 119px" bgcolor="#E2F3FC" class="style4">
										<font size="1" face="Verdana"><strong><%=getadminSmallCatLngStr("DtxtFunctions")%></strong></font></td>
										<td class="style3"><% HideFunctionTitle = True
										functionClass="TblFlowFunction" %><!--#include file="myFunctions.asp"--></td>
									</tr>
								</table>
								</td>
							</tr>
							</table>
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
		<table border="0" cellpadding="0" width="100%">
			<tr>
				<td width="77">
				<input type="submit" value="<%=getadminSmallCatLngStr("DtxtApply")%>" name="btnApply" class="OlkBtn"></td>
				<td width="77">
				<input type="submit" value="<%=getadminSmallCatLngStr("DtxtSave")%>" name="btnSave" class="OlkBtn"></td>
				<td><hr color="#0D85C6" size="1"></td>
				<td width="77">
				<input type="button" value="<%=getadminSmallCatLngStr("DtxtCancel")%>" name="B2" class="OlkBtn" onclick="javascript:if(confirm('<%=getadminSmallCatLngStr("DtxtConfCancel")%>'))window.location.href='adminSmallCat.asp?CatType=<%=CatType%>'"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
</table>
</form>
<script language="javascript">
NumUDAttach('frmEditSmallCat', 'Top', 'btnTopUp', 'btnTopDown');
</script>

<% End If %>
<!--#include file="bottom.asp" -->
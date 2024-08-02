<!--#include file="lang/adminSecRS.asp" -->
<!--#include file="myHTMLEncode.asp" -->
<html <% If Session("rtl") <> "" Then %>dir="rtl"<% End If %>>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">

<link rel="stylesheet" type="text/css" href="style/style_admin_<%=Session("style")%>.css"/>

<style type="text/css">
.style1 {
	font-weight: bold;
	background-color: #E1F3FD;
}
.style2 {
	background-color: #E1F3FD;
}
.style4 {
	font-family: Verdana;
	font-size: xx-small;
	color: #31659C;
}
</style>
</head>
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<% 
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="repVars.inc" -->
<body style="margin: 0">
<!--#include file="adminTradSubmit.asp"-->
<% If Request("SecID") <> "" Then %>
<% set rs = Server.CreateObject("ADODB.RecordSet")
sql = "select LineID, Name, Query from OLKSectionsRS where SecType = 'U' and SecID = " & Request("SecID")
rs.open sql, conn, 3, 1 %>
<table border="0" cellpadding="0" width="100%">
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table12">
			<tr>
				<td align="center" class="style1" style="width: 16px">
				&nbsp;</td>
				<td align="center" class="style1">
				<font size="1" face="Verdana" color="#31659C"><%=getadminSecRSLngStr("DtxtName")%>&nbsp;</font></td>
				<td align="center" class="style1">
				<font size="1" face="Verdana" color="#31659C"><%=getadminSecRSLngStr("DtxtQuery")%></font></td>
				<td align="center" class="style1" colspan="2">
				<font size="1" face="Verdana" color="#31659C"><%=getadminSecRSLngStr("LtxtTag")%></font></td>
				<td align="center" width="16" class="style2">&nbsp;</td>
			</tr>
			<%
			If rs.recordcount > 0 then
			do While NOT RS.EOF  %>
			<tr bgcolor="#F3FBFE">
			  <td valign="top" style="width: 16px; padding-top: 4px">
				<a href="adminSecRS.asp?SecID=<%=Request("SecID")%>&LineID=<%=rs("LineID")%>"><img border="0" src="images/<%=Session("rtl")%>flechaselec.gif"></a></td>
			  <td valign="top">
			  <font color="#31659C" face="Verdana" size="1"><%=rs("Name")%></font>
			  </td>
				<td valign="top" align="center">
				<img src="images/eye_icon.gif" dir="ltr" border="0" title="<%=Server.HTMLEncode(RS("Query"))%>"></td>
				<td valign="top" align="center">
				<font color="#31659C" face="Verdana" size="1">&lt;!--startRS<%=rs("LineID")%>--&gt;</font></td>
				<td valign="top" align="center">
				<font color="#31659C" face="Verdana" size="1">&lt;!--endRS<%=rs("LineID")%>--&gt;</font></td>
				<td valign="middle" width="16">
				<a href="javascript:if(confirm('<%=getadminSecRSLngStr("LtxtConfDel")%>'.replace('{0}', '<%=Replace(Server.HTMLEncode(Rs("Name")),"'","\'")%>')))window.location.href='adminSubmit.asp?cmd=del&SecID=<%=Request("SecID")%>&LineID=<%=rs("LineID")%>&submitCmd=adminSecRS';">
				<img border="0" src="images/remove.gif" width="16" height="16"></a></td>
				</tr>
				<% RS.MoveNext
				loop
				Else %>
				<tr>
					<td align="center" class="style1" colspan="6">
					<font size="1" face="Verdana" color="#31659C"><%=getadminSecRSLngStr("DtxtNoData")%></font></td>
				</tr>
				<% End If %>
		  </table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr>
				<td width="77">
				<input type="button" value="<%=getadminSecRSLngStr("DtxtNew")%>" name="btnNew" class="OlkBtn" onclick="javascript:window.location.href='adminSecRS.asp?SecID=<%=Request("SecID")%>&New=Y'"></td>
				<td><hr color="#0D85C6" size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
	<% If Request("New") = "Y" or Request("LineID") <> "" Then %>
	<tr>
	<form name="frmQuery" method="post" action="adminSubmit.asp">
	<input type="hidden" name="submitCmd" value="adminSecRS">
	<input type="hidden" name="cmd" value="edit">
	<input type="hidden" name="SecID" value="<%=Request("SecID")%>">
	<input type="hidden" name="LineID" value="<%=Request("LineID")%>">
		<td>
		<% If Request("LineID") <> "" Then
			sql = "select Name, Query from OLKSectionsRS where SecType = 'U' and SecID = " & Request("SecID") & " and LineID = " & Request("LineID") 
			set rs = conn.execute(sql) 
			Name = rs("Name")
			Query = rs("Query")
		Else
			Name = ""
			Query = "" %>
			<input type="hidden" name="queryDef" value="">
		<% End If %>
		<table border="0" cellpadding="0" width="100%">
			<tr>
				<td>
				<table border="0" cellpadding="0" width="100%">
					<tr>
						<td align="center" bgcolor="#E2F3FC"><b>
						<font size="1" face="Verdana" color="#31659C">
						<%=getadminSecRSLngStr("DtxtName")%>&nbsp;</font></b></td>
					</tr>
					<tr>
						<td valign="top" class="style3">
						<p align="center">
						<input name="Name" style="width: 100%; " class="input" value="<%=Server.HTMLEncode(Name)%>" size="50" maxlength="50">
						</td>
					</tr>
				</table>
				<table border="0" cellpadding="0" width="100%">
					<tr>
						<td valign="top">
						<table border="0" width="100%" cellpadding="0">
							<tr>
								<td valign="top" colspan="2" bgcolor="#E2F3FC">
								<table cellpadding="0" cellspacing="0" border="0" width="100%">
									<tr>
										<td rowspan="2">
											<textarea dir="ltr" rows="20" style="width: 100%" name="Query" cols="100" class="input" onkeypress="javascript:document.frmQuery.btnVerfyFilter.src='images/btnValidate.gif';document.frmQuery.btnVerfyFilter.style.cursor = 'hand';;document.frmQuery.valQuery.value='Y';"><%=myHTMLEncode(Query)%></textarea>
										</td>
										<td valign="top" width="1">
											<img style="cursor: hand; " src="images/qry_note.gif" id="btnNoteFilter" alt="<%=getadminSecRSLngStr("DtxtDefinition")%>" onclick="javascript:parent.doFldNote(19, 'Query', '<%=Request("SecID")%><%=Request("LineID")%>', <% If Request("LineID") <> "" Then %>null<% Else %>document.frmQuery.queryDef<% End If %>);">
										</td>
									</tr>
									<tr>
										<td valign="bottom" width="1">
											<img src="images/btnValidateDis.gif" id="btnVerfyFilter" alt="<%=getadminSecRSLngStr("DtxtValidate")%>" onclick="javascript:if (document.frmQuery.valQuery.value == 'Y')VerfyQuery();">
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
										<strong><%=getadminSecRSLngStr("DtxtVariables")%></strong></font></td>
										<td class="style3"><font face="Verdana" size="1" color="#4783C5"><span dir="ltr">@CardCode</span> = <%=getadminSecRSLngStr("DtxtClientCode")%><br>
								<span dir="ltr">@LanID</span> = <%=getadminSecRSLngStr("DtxtLanID")%></font></td>
									</tr>
									<tr>
										<td valign="top" style="width: 119px" bgcolor="#E2F3FC" class="style4">
										<font size="1" face="Verdana"><strong><%=getadminSecRSLngStr("DtxtFunctions")%></strong></font></td>
										<td class="style3"><% HideFunctionTitle = True
										functionClass="TblFlowFunction" %><!--#include file="myFunctions.asp"--></td>
									</tr>
								</table>
								</td>
							</tr>
							<tr>
								<td valign="top">
								<img src="images/spacer.gif"></td>
								<td>
								<img src="images/spacer.gif"></td>
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
				<input type="submit" value="<%=getadminSecRSLngStr("DtxtApply")%>" name="btnApply" class="OlkBtn"></td>
				<td width="77">
				<input type="submit" value="<%=getadminSecRSLngStr("DtxtSave")%>" name="btnSave" class="OlkBtn"></td>
				<td><hr color="#0D85C6" size="1"></td>
				<td width="77">
				<input type="button" value="<%=getadminSecRSLngStr("DtxtCancel")%>" name="B2" class="OlkBtn" onclick="javascript:if(confirm('<%=getadminSecRSLngStr("DtxtConfCancel")%>'))window.location.href='adminSecRS.asp?SecID=<%=Request("SecID")%>'"></td>
			</tr>
		</table>
		</td>
	</tr>
	</form>
	<% End If %>
</table>

<% If Request("New") = "Y" or Request("LineID") <> "" Then %>
<script type="text/javascript">
<!--
function VerfyQuery()
{
	document.frmVerfyQuery.Query.value = document.frmQuery.Query.value;
	document.frmVerfyQuery.submit();
}

function VerfyQueryVerified()
{
	document.frmQuery.btnVerfyFilter.src='images/btnValidateDis.gif'
	document.frmQuery.btnVerfyFilter.style.cursor = '';
	document.frmQuery.valQuery.value='N';
}
//-->
</script>
<form name="frmVerfyQuery" action="verfyQuery.asp" method="post" target="iVerfyQuery">
	<iframe id="iVerfyQuery" name="iVerfyQuery" style="display: none" src=""></iframe>
	<input type="hidden" name="type" value="secSubRS">
	<input type="hidden" name="Query" value="">
	<input type="hidden" name="parent" value="Y">
</form>
<% End If %>
<% Else %>
<table border="0" cellpadding="0" width="100%">
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table12">
			<tr>
				<td align="center" class="style1"><%=getadminSecRSLngStr("LtxtSaveSec")%></td>
			</tr>
		</table>
		</td>
	</tr>
</table>
<% End If %>
</body>

</html>

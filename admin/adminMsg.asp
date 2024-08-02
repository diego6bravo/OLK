<!--#include file="top.asp" -->
<!--#include file="lang/adminMsg.asp" -->

<head>
<% conn.execute("use [" & Session("OLKDB") & "]") %>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style type="text/css">
.style1 {
	background-color: #F7FBFF;
}
</style>
</head>
<% If Session("style") = "ie" Then %>
<br>
<% End If %>
<!--#include file="adminTradSubmit.asp"-->
<table border="0" cellpadding="0" width="100%" id="table3">
	<tr>
		<td bgcolor="#E1F3FD"><b><font color="#FFFFFF">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><%=getadminMsgLngStr("LttlEMails")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif">
		<font face="Verdana" size="1" color="#4783C5"><%=getadminMsgLngStr("LttlEMailsNote")%></font></td>
	</tr>
	<% If Request("MsgID") = "" Then %>
	<tr>
		<td width="100%">
		<table border="0" cellpadding="0" width="100%" id="table6">
			<tr>
				<td colspan="2">
				<table border="0" cellpadding="0" id="table12">
					<tr>
						<td align="center" bgcolor="#E2F3FC"></td>
						<td align="center" bgcolor="#E2F3FC"><font size="1" face="Verdana" color="#31659C">
						<b><%=getadminMsgLngStr("DtxtType")%></b>&nbsp;</font></td>
						<td align="center" bgcolor="#E2F3FC" colspan="2"><font size="1" face="Verdana" color="#31659C">
						<b><%=getadminMsgLngStr("LtxtSubject")%></b>&nbsp;</font></td>
					</tr>
					<% 
					sql = "select T0.MsgID, IsNull(T1.AlterMsgName, T0.MsgName) MsgName, T0.Header " & _
							"from OLKAutMsg T0 " & _
							"left outer join OLKAutMsgAlterNames T1 on T1.MsgID = T0.MsgID and T1.LanID = " & Session("LanID") & " " & _
							"order by 2 asc "
					set rs = conn.execute(sql)
					do while not rs.eof %>
					<tr>
						<td bgcolor="#F3FBFE">
						<a href="adminMsg.asp?MsgID=<%=rs("MsgID")%>"><img border="0" src="images/<%=Session("rtl")%>flechaselec.gif"></a></td>
						<td bgcolor="#F3FBFE"><font size="1" face="Verdana" color="#4783C5"><%=rs("MsgName")%>&nbsp;</font></td>
						<td bgcolor="#F3FBFE"><span dir="ltr"><font size="1" face="Verdana" color="#4783C5"><%=rs("Header")%>&nbsp;</font></span>
						</td>
						<td bgcolor="#F3FBFE" style="width: 16px">
						<a href="javascript:doFldTrad('AutMsg', 'MsgID', '<%=rs("MsgID")%>', 'AlterHeader', 'M', null);"><img src="images/trad.gif" alt="|D:txtTranslate|" border="0"></a></td>
					</tr>
					<% rs.movenext
					loop %>
				</table>
				</td>
			</tr>
			</table>
		</td>
	</tr>
	<% ElseIf Request("MsgID") <> "" Then
	sql = "select IsNull(T1.AlterMsgName, T0.MsgName) MsgName, T0.Query, T0.Header, T0.Message " & _
	"from OLKAutMsg T0 " & _
	"left outer join OLKAutMsgAlterNames T1 on T1.MsgID = T0.MsgID and T1.LanID = " & Session("LanID") & " " & _
	"where T0.MsgID = " & Request("MsgID")
	set rs = conn.execute(sql) %>
<form method="POST" action="adminSubmit.asp" name="frmEditMsg" onsubmit="javascript:return valFrm()">
	<tr>
		<td bgcolor="#E1F3FD">&nbsp;<b><font face="Verdana" size="1" color="#31659C"><%=getadminMsgLngStr("LttlEditMsg")%> 
		- <%=rs("MsgName")%></font></b></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table20">
			<tr>
				<td valign="top">
				<table border="0" width="100%" id="table23" cellpadding="0">
					<tr>
						<td valign="top" bgcolor="#E2F3FC" style="width: 100px">
						<font face="Verdana" size="1" color="#31659C">
						<strong><%=getadminMsgLngStr("DtxtQuery")%></strong></font></td>
						<td class="style1">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td rowspan="2">
									<textarea rows="10" dir="ltr" name="Query" cols="87" class="input" onkeypress="javascript:document.frmEditMsg.btnVerfyFilter.src='images/btnValidate.gif';document.frmEditMsg.btnVerfyFilter.style.cursor = 'hand';;document.frmEditMsg.valQuery.value='Y';"><%=myHTMLEncode(rs("Query"))%></textarea>
								</td>
								<td valign="top">
									<img style="cursor: hand; " src="images/qry_note.gif" id="btnNoteFilter" alt="<%=getadminMsgLngStr("DtxtDefinition")%>" onclick="javascript:doFldNote(5, 'Query', <%=Request("MsgID")%>, null);">
								</td>
							</tr>
							<tr>
								<td valign="bottom">
									<img src="images/btnValidateDis.gif" id="btnVerfyFilter" alt="<%=getadminMsgLngStr("DtxtValidate")%>" onclick="javascript:if (document.frmEditMsg.valQuery.value == 'Y')VerfyQuery();">
									<input type="hidden" name="valQuery" value="N">	
							</td>
							</tr>
						</table>
						</td>
					</tr>
					<tr>
						<td valign="top" bgcolor="#E2F3FC" style="width: 100px">
						<font face="Verdana" size="1" color="#31659C">
						<strong><%=getadminMsgLngStr("DtxtVariables")%></strong></font></td>
						<td class="style1">
						<font size="1" color="#4783C5" face="Verdana">
						<span dir="ltr">@CardCode</span> = <%=getadminMsgLngStr("DtxtClientCode")%><br>
						<span dir="ltr">@txtLanID</span> = <%=getadminMsgLngStr("DtxtLanID")%></font></td>
					</tr>
					<tr>
						<td valign="top" bgcolor="#E2F3FC" style="width: 100px">
						<font face="Verdana" size="1" color="#31659C">
						<strong><%=getadminMsgLngStr("DtxtFunctions")%></strong></font></td>
						<td class="style1">
						<% HideFunctionTitle = True
						functionClass="TblFlowFunction" %><!--#include file="myFunctions.asp"--></td>
					</tr>
					<tr>
						<td valign="top" bgcolor="#E2F3FC" style="width: 100px">
						<font face="Verdana" size="1" color="#31659C">
						<strong><%=getadminMsgLngStr("LtxtSubject")%></strong></font></td>
						<td valign="top" class="style1">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td valign="top"><textarea rows="3" name="Header" style="width: 420px;" cols="67" class="input" onkeydown="return chkMax(event, this, 100);"><%=myHTMLEncode(rs("Header"))%></textarea>
								</td>
								<td valign="top" width="16" style="padding-top: 25px;"><a href="javascript:doFldTrad('AutMsg', 'MsgID', '<%=Request("MsgID")%>', 'AlterHeader', 'M', null);"><img src="images/trad.gif" alt="|D:txtTranslate|" border="0"></a>
								</td>
								<td><select size="10" name="HeaderQryFields" style="width:120px; height:126px; " ondblclick="javascript:document.frmEditMsg.Header.value+=this.value;">
									<% If rs("Query") <> "" Then
										sql = "declare @CardCode nvarchar(15) set @CardCode = '' declare @LanID int set @LanID = -1 " & rs("Query")
										set rd = conn.execute(sql)
										For each item in rd.Fields
										If item.Name <> "" Then %>
										<option value="{<%=myHTMLEncode(item.Name)%>}"><%=myHTMLEncode(item.Name)%></option>
										<% End If
										next
									End If %>
									</select>
								</td>
							</tr>
						</table>
						</td>
					</tr>
					<tr>
						<td valign="top" bgcolor="#E2F3FC" style="width: 100px">
						<font face="Verdana" size="1" color="#31659C"><strong><%=getadminMsgLngStr("DtxtMessage")%></strong></font></td>
						<td valign="top" class="style1">
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td>
								<table cellpadding="0" cellspacing="0" border="0">
									<tr>
										<td><textarea rows="10" style="width: 420px;" name="Message" cols="67" class="input"><%=myHTMLEncode(rs("Message"))%></textarea>
										</td>
										<td valign="bottom" width="16"><a href="javascript:doFldTrad('AutMsg', 'MsgID', '<%=Request("MsgID")%>', 'AlterMessage', 'M', null);"><img src="images/trad.gif" alt="|D:txtTranslate|" border="0"></a>
										</td>
									</tr>
								</table>
								</td>
								<td>
								<select size="10" name="MsgQryFields" style="width:120px; height:126px" ondblclick="javascript:document.frmEditMsg.Message.value+=this.value;">
								<% If rs("Query") <> "" Then
									For each item in rd.Fields
									If item.Name <> "" Then %>
									<option value="{<%=myHTMLEncode(item.Name)%>}"><%=myHTMLEncode(item.Name)%></option>
									<% End If
									next
								End If %>
								</select>
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
		<table border="0" cellpadding="0" width="100%" id="table5">
			<tr>
				<td width="77">
				<input type="submit" value="<%=getadminMsgLngStr("DtxtApply")%>" name="btnApply" class="OlkBtn"></td>
				<td width="77">
				<input type="submit" value="<%=getadminMsgLngStr("DtxtSave")%>" name="btnSave" class="OlkBtn"></td>
				<td><hr color="#0D85C6" size="1"></td>
				<td width="77">
				<input type="button" value="<%=getadminMsgLngStr("DtxtCancel")%>" name="B2" class="OlkBtn" onclick="if(confirm('<%=getadminMsgLngStr("DtxtConfCancel")%>'))window.location.href='adminMsg.asp';"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
	<input type="hidden" name="submitCmd" value="adminMsg">
	<input type="hidden" name="MsgID" value="<%=Request("MsgID")%>">
</form>
<script language="javascript">
function valFrm()
{
	if (document.frmEditMsg.Query.value != '' && document.frmEditMsg.valQuery.value == 'Y')
	{
		alert('<%=getadminMsgLngStr("LtxtValQryVal")%>');
		document.frmEditMsg.btnVerfyFilter.focus();
		return false;
	}
	else if (document.frmEditMsg.Header.value == '')
	{
		alert('<%=getadminMsgLngStr("LtxtValEmlSubject")%>');
		document.frmEditMsg.Header.focus();
		return false;
	}
	else if (document.frmEditMsg.Message.value == '')
	{
		alert('<%=getadminMsgLngStr("LtxtValEmailMsg")%>');
		document.frmEditMsg.Message.focus();
		return false;
	}
	return true;
}
function VerfyQuery()
{
	document.frmVerfyQuery.Query.value = document.frmEditMsg.Query.value;
	if (document.frmVerfyQuery.Query.value != '')
	{
		document.frmVerfyQuery.submit();
	}
	else
	{
		VerfyQueryVerified();
		for (var i = document.frmEditMsg.HeaderQryFields.length-1;i>=0;i--)
		{
			document.frmEditMsg.HeaderQryFields.remove(i);
			document.frmEditMsg.MsgQryFields.remove(i);
		}
	}
}
function VerfyQueryVerified()
{
	document.frmEditMsg.btnVerfyFilter.src='images/btnValidateDis.gif'
	document.frmEditMsg.btnVerfyFilter.style.cursor = '';
	document.frmEditMsg.valQuery.value='N';
}

function getHeaderQueryFields() { return document.frmEditMsg.HeaderQryFields; }
function getMsgQueryFields() { return document.frmEditMsg.MsgQryFields; }
</script>
<% Else %>
	<% End If %>
	</table>
<form name="frmVerfyQuery" action="verfyQuery.asp" method="post" target="iVerfyQuery">
	<iframe id="iVerfyQuery" name="iVerfyQuery" style="display: none" src=""></iframe>
	<input type="hidden" name="type" value="MailMsg">
	<input type="hidden" name="Query" value="">
	<input type="hidden" name="by" value="">
	<input type="hidden" name="parent" value="Y">
</form>

<!--#include file="bottom.asp" -->
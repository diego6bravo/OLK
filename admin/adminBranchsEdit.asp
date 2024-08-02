<!--#include file="top.asp" -->
<!--#include file="lang/adminBranchsEdit.asp" -->
<!--#include file="adminTradSubmit.asp"-->
<!--#include file="accountControl.asp"-->  

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<script language="javascript">
<!-- Start BEGIN
var fldAcctCode;
var fldAcctName;
function getCuenta(AcctCode, AcctName, Type)
{
	fldAcctCode = AcctCode;
	frlAcctName = AcctName;
	Start2('cuentas.asp', 500, 200, Type);
}

function Start2(theURL, popW, popH, type) { // V 1.0
var winleft = (screen.width - popW) / 2;
var winUp = (screen.height - popH) / 2;
winProp = 'width='+popW+',height='+popH+',left='+winleft+',top='+winUp+',toolbar=no,scrollbars=yes,menubar=no,location=no,resizable=no'
theURL2 = theURL+'?update='+type
OpenWin = window.open(theURL2, "CtrlWindow2", winProp)

}
// Start END -->
function setCuenta(AcctCode, AcctName, Update)
{
	fldAcctCode.value = AcctCode;
	frlAcctName.value = AcctName;
}
</script>
<% 
conn.execute("use [" & Session("OLKDB") & "]")
GetAdminQuery rs, 3, "D", Request("branchIndex")

set rd = server.createobject("ADODB.RecordSet")
GetQuery rd, 1, null, null
%>
<script language="javascript">
function valFrm()
{
	if (document.frmEditBranch.branchName.value == '')
	{
		alert("<%=getadminBranchsEditLngStr("LtxtValBranchNam")%>");
		document.frmEditBranch.branchName.focus();
		return false;
	}
	return true;
}
</script>
</head>

<form method="POST" action="submitBranchs.asp" name="frmEditBranch" onsubmit="return valFrm();">
<table border="0" cellpadding="0" width="100%" id="table6">
	<tr>
		<td bgcolor="#E1F3FD"><b><font size="1" color="#31659C" face="Verdana">
		<%=getadminBranchsEditLngStr("LttlEditBranch")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif">&nbsp;<font color="#4783C5" face="Verdana" size="1"><%=getadminBranchsEditLngStr("LttlEditBranchNote")%></font></td>
	</tr>
	<tr>
		<td bgcolor="#FFFFFF">
		<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table9">
			<tr bgcolor="#F5FBFE">
				<td><font face="Verdana" size="1" color="#4783C5">
				<%=getadminBranchsEditLngStr("DtxtName")%></font></td>
				<td><table cellpadding="0" cellspacing="0" border="0">
				<tr>
					<td style="width: 400px">
					<input type="text" name="branchName" class="input" size="30" style="width: 100%; " value="<% If Not IsNull(rs("branchName")) Then %><%=Server.HTMLEncode(rs("branchName"))%><% End If %>" onkeydown="return chkMax(event, this, 50);">
					</td>
					<td width="16" valign="bottom">
					<a href="javascript:doFldTrad('Branchs', 'branchIndex', <%=Request("branchIndex")%>, 'alterBranchName', 'T', null);"><img src="images/trad.gif" alt="<%=getadminBranchsEditLngStr("DtxtTranslate")%>" border="0"></a>
					</td>
					<td><font face="Verdana" size="1"><font size="1">
		<input <%If rs("Active") = "Y" then response.write "checked"%> type="checkbox" name="Active" value="Y" id="fp1" class="noborder"></font><font color="#4783C5"><label for="fp1"><%=getadminBranchsEditLngStr("DtxtActive")%></label></font></font></td>
				</tr>
			</table></td>
			</tr>
			<tr bgcolor="#F5FBFE">
				<td><font face="Verdana" size="1" color="#4783C5">
				<%=getadminBranchsEditLngStr("DtxtWarehouse")%></font></td>
				<td>
		<font size="1" face="Verdana">
		<select size="1" name="whsCode" class="input">
		<% do while not rd.eof %>
		<option <% If rd("whsCode") = rs("whsCode") then response.write "selected"%> value="<%=rd("whsCode")%>"><%=myHTMLEncode(rd("whsName"))%></option>
		<% rd.movenext
		loop %>
		</select></font></td>
			</tr>
			<tr bgcolor="#F5FBFE">
				<td><font face="Verdana" size="1" color="#4783C5">
				<%=getadminBranchsEditLngStr("LtxtOQUTSeries")%></font></td>
				<td><font size="1" face="Verdana"><select size="1" name="OQUTSeries" class="input">
				<option value=""><%=getadminBranchsEditLngStr("DtxtDefault")%>&nbsp;<%=getadminBranchsEditLngStr("DtxtOLK")%></option>
				<%  
				GetQuery rd, 4, 23, null
				do While NOT RD.EOF %>
				<option <% If rd("Series") = rs("OQUTSeries") Then %>selected<%end if%> value="<%=rd("Series")%>"><%=myHTMLEncode(rd("SeriesName"))%></option>
				<%  RD.MoveNext
				loop %>
				</select></font></td>
			</tr>
			<tr bgcolor="#F5FBFE">
				<td><font face="Verdana" size="1" color="#4783C5">
				<%=getadminBranchsEditLngStr("LtxtORDRSeries")%></font></td>
				<td><font size="1" face="Verdana"><select size="1" name="ORDRSeries" class="input">
				<option value=""><%=getadminBranchsEditLngStr("DtxtDefault")%>&nbsp;<%=getadminBranchsEditLngStr("DtxtOLK")%></option>
				<%  
				GetQuery rd, 4, 17, null
				do While NOT RD.EOF %>
				<option <% If rd("Series") = rs("ORDRSeries") Then %>selected<%end if%> value="<%=rd("Series")%>"><%=myHTMLEncode(rd("SeriesName"))%></option>
				<%  RD.MoveNext
				loop %>
				</select></font></td>
			</tr>
			<tr bgcolor="#F5FBFE">
				<td><font face="Verdana" size="1" color="#4783C5">
				<%=getadminBranchsEditLngStr("LtxtODLNSeries")%></font></td>
				<td><font size="1" face="Verdana">
				<select size="1" name="ODLNSeries" class="input">
				<option value=""><%=getadminBranchsEditLngStr("DtxtDefault")%>&nbsp;<%=getadminBranchsEditLngStr("DtxtOLK")%></option>
				<% 
				GetQuery rd, 4, 15, null
				do While NOT RD.EOF %>
				<option <% If rd("Series") = rs("ODLNSeries") Then %>selected<%end if%> value="<%=rd("Series")%>"><%=myHTMLEncode(rd("SeriesName"))%></option>
				<%  RD.MoveNext
				loop %>
				</select></font></td>
			</tr>
			<tr bgcolor="#F5FBFE">
				<td><font face="Verdana" size="1" color="#4783C5">
				<%=getadminBranchsEditLngStr("LtxtOINVSeries")%></font></td>
				<td><font size="1" face="Verdana"><select size="1" name="OINVSeries" class="input">
				<option value=""><%=getadminBranchsEditLngStr("DtxtDefault")%>&nbsp;<%=getadminBranchsEditLngStr("DtxtOLK")%></option>
				<%  
				GetQuery rd, 4, 13, null
				do While NOT RD.EOF %>
				<option <% If rd("Series") = rs("OINVSeries") Then %>selected<%end if%> value="<%=rd("Series")%>"><%=myHTMLEncode(rd("SeriesName"))%></option>
				<% RD.MoveNext
				loop %>
				</select></font></td>
			</tr>
			<tr bgcolor="#F5FBFE">
				<td><font face="Verdana" size="1" color="#4783C5">
				<%=getadminBranchsEditLngStr("LtxtOINVResSeries")%></font></td>
				<td><font size="1" face="Verdana"><select size="1" name="OINVResSeries" class="input">
				<option value=""><%=getadminBranchsEditLngStr("DtxtDefault")%>&nbsp;<%=getadminBranchsEditLngStr("DtxtOLK")%></option>
				<%  
				GetQuery rd, 4, 13, null
				rd.movefirst
				do While NOT RD.EOF %>
				<option <% If rd("Series") = rs("OINVResSeries") Then %>selected<%end if%> value="<%=rd("Series")%>"><%=myHTMLEncode(rd("SeriesName"))%></option>
				<% RD.MoveNext
				loop %>
				</select></font></td>
			</tr>
			<tr bgcolor="#F5FBFE">
				<td><font face="Verdana" size="1" color="#4783C5">
				<%=getadminBranchsEditLngStr("LtxtORCTSeries")%></font></td>
				<td><font size="1" face="Verdana"><select size="1" name="ORCTSeries" class="input">
				<option value=""><%=getadminBranchsEditLngStr("DtxtDefault")%>&nbsp;<%=getadminBranchsEditLngStr("DtxtOLK")%></option>
				<%  
				GetQuery rd, 4, 24, null
				do While NOT RD.EOF %>
				<option <% If rd("Series") = rs("ORCTSeries") Then %>selected<%end if%> value="<%=rd("Series")%>"><%=myHTMLEncode(rd("SeriesName"))%></option>
				<% RD.MoveNext
				loop %>
				</select></font></td>
			</tr>
			<tr bgcolor="#F5FBFE">
				<td><font face="Verdana" color="#4783c5" size="1">
				<%=getadminBranchsEditLngStr("LtxtOINVSeries")%> - (<%=getadminBranchsEditLngStr("LtxtInvRec")%>)</font></td>
				<td><font size="1" face="Verdana">
				<select size="1" name="OIRISeries" class="input">
				<option value=""><%=getadminBranchsEditLngStr("DtxtDefault")%>&nbsp;<%=getadminBranchsEditLngStr("DtxtOLK")%></option>
				<% 
				GetQuery rd, 4, 13, null
				do While NOT RD.EOF %>
				<option <% If rd("Series") = rs("OIRISeries") Then %>selected<%end if%> value="<%=rd("Series")%>"><%=myHTMLEncode(rd("SeriesName"))%></option>
				<% RD.MoveNext
				loop %>
				</select></font></td>
			</tr>
			<tr bgcolor="#F5FBFE">
				<td><font face="Verdana" color="#4783c5" size="1">
				<%=getadminBranchsEditLngStr("LtxtORCTSeries")%> - (<%=getadminBranchsEditLngStr("LtxtInvRec")%>)</font></td>
				<td><font size="1" face="Verdana">
				<select size="1" name="OIRRSeries" class="input">
				<option value=""><%=getadminBranchsEditLngStr("DtxtDefault")%>&nbsp;<%=getadminBranchsEditLngStr("DtxtOLK")%></option>
				<%
				GetQuery rd, 4, 24, null
				do While NOT RD.EOF %>
				<option <% If rd("Series") = rs("OIRRSeries") Then %>selected<%end if%> value="<%=rd("Series")%>"><%=myHTMLEncode(rd("SeriesName"))%></option>
				<% RD.MoveNext
				loop %>
				</select></font></td>
			</tr>
			<tr bgcolor="#F5FBFE">
				<td><font face="Verdana" size="1" color="#4783C5">
				<%=getadminBranchsEditLngStr("LtxtCashAcct")%></font></td>
				<td>
				<% 
				Dim myAccount
				set myAccount = New AccountControl
				myAccount.ID = "CashAcct"
				myAccount.Value = rs("CashAcct")
				myAccount.DisplayValue = rs("CashAcctDisp")
				myAccount.Description = rs("CashAcctName")
				myAccount.AccountType = "cash"
				myAccount.GenerateAccount %></td>
			</tr>
			<tr bgcolor="#F5FBFE">
				<td><font face="Verdana" color="#4783c5" size="1">
				<%=getadminBranchsEditLngStr("LtxtCheckAcct")%></font></td>
				<td><%
				myAccount.ID = "CheckAcct"
				myAccount.Value = rs("CheckAcct")
				myAccount.DisplayValue = rs("CheckAcctDisp")
				myAccount.Description = rs("CheckAcctName")
				myAccount.AccountType = "check"
				myAccount.GenerateAccount %></td>
			</tr>
			<% 
			GetAdminQuery rd, 3, "CCA", Request("branchIndex")
			do while not rd.eof %>
			<tr bgcolor="#F5FBFE">
				<td>
				<font face="Verdana" color="#4783c5" size="1">
				<%=getadminBranchsEditLngStr("LtxtCredAcct")%> - <%=rd("CardName")%></font></td>
				<td>
				<%
				myAccount.ID = "CreditAcctCode" & rd("CreditCard")
				myAccount.Value = rd("AcctCode")
				myAccount.DisplayValue = rd("AcctDisp")
				myAccount.Description = rd("AcctName")
				myAccount.AccountType = "check"
				myAccount.GenerateAccount %></td>
			</tr>
			<% rd.movenext
			loop  %>
			<tr bgcolor="#F5FBFE">
				<td><font face="Verdana" color="#4783c5" size="1">
				<%=getadminBranchsEditLngStr("LtxtCashAcct")%> - (<%=getadminBranchsEditLngStr("LtxtInvRec")%>)</font></td>
				<td><%
				myAccount.ID = "OIRCashAcct"
				myAccount.Value = rs("OIRCashAcct")
				myAccount.DisplayValue = rs("OIRCashAcctDisp")
				myAccount.Description = rs("OIRCashAcctName")
				myAccount.AccountType = "cash"
				myAccount.GenerateAccount %></td>
			</tr>
			<tr bgcolor="#F5FBFE">
				<td><font face="Verdana" color="#4783c5" size="1">
				<%=getadminBranchsEditLngStr("LtxtCheckAcct")%> - (<%=getadminBranchsEditLngStr("LtxtInvRec")%>)</font></td>
				<td><%
				myAccount.ID = "OIRCheckAcct"
				myAccount.Value = rs("OIRCheckAcct")
				myAccount.DisplayValue = rs("OIRCheckAcctDisp")
				myAccount.Description = rs("OIRCheckAcctName")
				myAccount.AccountType = "check"
				myAccount.GenerateAccount %></td>
			</tr>
			<%
			GetAdminQuery rd, 3, "CCIRA", Request("branchIndex")
			do while not rd.eof %>
			<tr bgcolor="#F5FBFE">
				<td>
				<p>
				<font face="Verdana" color="#4783c5" size="1">
				<%=getadminBranchsEditLngStr("LtxtCredAcct")%> - <%=rd("CardName")%> (<%=getadminBranchsEditLngStr("LtxtInvRec")%>)</font></td>
				<td height="18"><%
				myAccount.ID = "OIRCreditAcctCode" & rd("CreditCard")
				myAccount.Value = rd("AcctCode")
				myAccount.DisplayValue = rd("AcctDisp")
				myAccount.Description = rd("AcctName")
				myAccount.AccountType = "check"
				myAccount.GenerateAccount %></td>
			</tr>
			<% rd.movenext
			loop  %>
		</table>
		</td>
	</tr>
	<tr bgcolor="#F5FBFE">
		<td>
		<table border="0" cellpadding="0" width="100%" id="table7">
			<tr>
				<td width="77">
				<input type="submit" value="<%=getadminBranchsEditLngStr("DtxtApply")%>" name="btnApply" class="OlkBtn"></td>

				<td width="77">
				<input type="submit" value="<%=getadminBranchsEditLngStr("DtxtSave")%>" name="btnSave" class="OlkBtn"></td>
				<td><hr></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td><font face="Verdana" size="1">&nbsp;
		</font>
		</td>
	</tr>
</table>
	<input type="hidden" name="cmd" value="updateBranch">
	<input type="hidden" name="redir" value="adminBranchsEdit.asp">
	<input type="hidden" name="branchIndex" value="<%=Request("branchIndex")%>">
</form>
<script language="javascript" src="accountControl.js"></script>
<!--#include file="bottom.asp" -->
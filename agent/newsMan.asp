<!--#include file="clientInc.asp"-->
<!--#include file="agentTop.asp"-->
<% If userType <> "V" Then Response.Redirect "default.asp" %>
<% If Not (myAut.HasAuthorization(4)) Then Response.Redirect "unauthorized.asp" %>
<head>
<style type="text/css">
.style1 {
	text-align: center;
}
</style>
</head>

<% addLngPathStr = "" %>
<!--#include file="lang/newsMan.asp" -->
<!--#include file="genman/adminTradForm.asp"-->
<% 
sql = "select newsIndex, newsDate, newsTitle, Case Status When 'A' Then N'" & getnewsManLngStr("DtxtYes") & "' When 'N' Then N'" & getnewsManLngStr("DtxtNo") & "' End StatusStr from olknews where Status <> 'D' order by newsdate desc"
set rs = conn.execute(sql)
%>
<script type="text/javascript">
function valFrm()
{
	var found = false;
	if (document.frmNews.delID)
	{
		if (document.frmNews.delID.length)
		{
			for (var i = 0;i<document.frmNews.delID.length;i++)
			{
				if (document.frmNews.delID[i].checked)
				{
					found = true;
					break;
				}
			}
		}
		else
		{
			found = document.frmNews.delID.checked;
		}
	}
	
	if (!found)
	{
		alert('<%=getnewsManLngStr("LtxtValSelNews")%>');
		return false;
	}
	else
	{
		return confirm('<%=getnewsManLngStr("LtxtConfDelNews")%>');
	}
}
</script>
<form name="frmNews" action="genman/newsSubmit.asp" method="post" onsubmit="return valFrm();">
<input type="hidden" name="submitCmd" value="del">
<table border="0" cellpadding="0" width="100%">
	<tr>
		<td class="GeneralTlt"><%=getnewsManLngStr("LttlNews")%></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr class="FirmTlt3">
				<td align="center"style="width: 16px">
				&nbsp;</td>
				<td align="center"style="width: 16px">
				&nbsp;</td>
				<td style="width: 100px" align="center">
				<nobr><%=getnewsManLngStr("DtxtDateOfPub")%>&nbsp;</nobr></td>
				<td align="center">
				<%=getnewsManLngStr("DtxtTitle")%></td>
				<td style="width: 100px" align="center">
				<%=getnewsManLngStr("DtxtActive")%></td>
			</tr>
			<% do while not rs.eof %>
			<tr class="GeneralTbl">
				<td style="width: 16px" class="style1">
				<img src="images/checkbox_off.jpg" border="0" onclick="doCheckDel(this, <%=rs("newsIndex")%>);">
				<input type="checkbox" name="delID" id="delID<%=rs("newsIndex")%>" value="<%=rs("newsIndex")%>" style="display: none;"></td>
				<td style="width: 16px">
				<a href="javascript:doMyLink('newsEdit.asp', 'newsIndex=<%=rs("newsIndex")%>', '_self');">
				<img border="0" src="images/<%=Session("rtl")%>flechaselec.gif" width="15" height="13"></a></td>
				<td style="width: 100px"><%=FormatDate(rs("newsDate"), True)%>&nbsp;</td>
				<td>
				<table cellpadding="0" cellspacing="0" border="0" width="100%">
					<tr class="GeneralTbl">
						<td valign="middle"><%=rs("newsTitle")%>
						</td>
						<td width="16">
						<a href="javascript:doFldTrad('News', 'newsIndex', <%=rs("newsIndex")%>, 'alterNewsTitle', 'T', null);"><img src="images/trad.gif" alt="<%=getnewsManLngStr("DtxtTranslate")%>" border="0"></a></td>
					</tr>
				</table>
				</td>
				<td style="width: 100px; text-align: center; ">
				<%=rs("StatusStr")%></td>
			</tr>
			<% rs.movenext
			loop %>
		</table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table5">
			<tr>
				<td width="77">
				<input type="submit" value="<%=getnewsManLngStr("DtxtDelete")%>" name="btnDelete"></td>
				<td>&nbsp;</td>
				<td width="77">
				<input type="button" value="<%=getnewsManLngStr("LtxtNewNews")%>" name="btnNew" onclick="javascript:doMyLink('newsEdit.asp', '', '_self');"></td>
			</tr>
		</table>
		</td>
	</tr>
</table>
</form>
<script type="text/javascript">
<!--
function doCheckDel(Img, LogNum)
{
	if (!document.getElementById('delID' + LogNum).checked)
	{
		document.getElementById('delID' + LogNum).checked = true;
		Img.src = 'images/checkbox_on.jpg';
	}
	else
	{
		document.getElementById('delID' + LogNum).checked = false;
		Img.src = 'images/checkbox_off.jpg';
	}
}
//-->
</script>
<!--#include file="agentBottom.asp"-->
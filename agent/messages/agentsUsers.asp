<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="../chkLogin.asp" -->
<!--#include file="lang/agentsUsers.asp" -->
<html <% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>>
<!--#include file="../myHTMLEncode.asp"-->
<head>
<%
set rx = Server.CreateObject("ADODB.recordset")
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKMessageAgents" & Session("ID")
cmd.Parameters.Refresh()
cmd("@LanID") = Session("LanID")
cmd("@SlpCode") = Session("vendid")
If Request("agentsusers") <> "" Then cmd("@Checked") = Request("agentsusers")
rx.open cmd, , 3, 1
%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<% If Request.Form("btnAccept") <> "" Then %>
<script language="javascript" src="../general.js"></script>
<script type="text/javascript">
opener.agentsTo("<%=Request("agentsusers")%>");
window.close();
</script>
<% End IF %>
<script type="text/javascript">
var checkflag = "false";
function check(field) 
{
	All = field.checked;
	agentsusers = document.form1.agentsusers;
	for (var i = 0;i<agentsusers.length;i++)
	{
		agentsusers[i].checked = All;
	}
}

function checkAll()
{
var All = true;
<% If rx.recordcount > 1 Then %>
var agentsusers = document.form1.agentsusers;
for (var i = 0;i<agentsusers.length;i++)
{
	if (!agentsusers[i].checked)
	{
		All = false;
		break;
	}
}
<% Else %>
if (!document.form1.agentsusers.checked) { All = false; }
<% End If %>
if (document.form1.C1 != null) document.form1.C1.checked = All;
}

</script>
<link rel="stylesheet" type="text/css" href="../design/<%=GetSelDes%>/style/stylePopUp.css">
<title><%=getagentsUsersLngStr("LttlOLKUsers")%></title>
</head>
<!--#include file="../design/popvars.inc" -->
<body topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0">
<form method="post" action="agentsUsers.asp" name="form1">
            <table border="0" cellpadding="0" width="100%" id="table1">
				<% If tblCustTtl = "" Then %>
              <tr class="GeneralTlt">
                <td id="tdMyTtl" width="50%"><%=getagentsUsersLngStr("LttlOLKUsers")%>:</td>
              </tr><% Else %>
				<% AddPath = "../" %>
				<%=Replace(Replace(tblCustTtl, "{txtTitle}", getagentsUsersLngStr("LttlOLKUsers")), "{AddPath}", "../")%>
				<% End If %>
              <tr>
                <td width="50%" height="42">
                <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="table2">
              <% do while not rx.eof %>
                  <tr class="GeneralTbl">
                    <td width="100%">
				<input type="checkbox" style="border-style:solid; border-width:0; background:background-image" name="agentsusers" value="<%=myHTMLEncode(RX("U_Name"))%>" id="fps<%=RX("User_Code")%>" <%=RX("checked")%> onclick="javascript:checkAll()"><label for="fps<%=RX("User_Code")%>"><%=RX("U_Name")%></label></td>
                  </tr>
        <% rx.movenext
        loop %>
        <% If rx.recordcount > 1 then %>
                  <tr class="GeneralTbl">
                    <td width="100%">
					<input type="checkbox" <% If UBound(Split(Request("agentsusers"),", ")) = rx.recordcount-1 then Response.Write "checked" %> style="border-style:solid; border-width:0; background:background-image" name="C1" value="ON" onclick="check(this)" id="fp1"><label for="fp1"><%=getagentsUsersLngStr("DtxtAll")%></label></td>
                  </tr>
        <% End If %>
                </table>
                </td>
              </tr>
              </table>
            <center><input type="submit" value="<%=getagentsUsersLngStr("DtxtAccept")%>" name="btnAccept"></center>
			<input type="hidden" name="AddPath" value="../">
			<input type="hidden" name="pop" value="Y">

<% If setCustTtl and userType = "C" Then %>
<script language="javascript" src="../setTltBg.js.asp?custTtlBgL=<%=custTtlBgL%>&custTtlBgM=<%=custTtlBgM%>&AddPath=../"></script>
<script language="javascript">setTtlBg(false);</script>
<% End If %></body></html><% 
conn.close
set rx = nothing
%>
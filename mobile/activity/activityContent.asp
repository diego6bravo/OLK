<% addLngPathStr = "activity/" %>
<!--#include file="lang/activityContent.asp" -->
<head>
<style type="text/css">
.style1 {
				font-family: Verdana;
				font-size: xx-small;
}
.style2 {
				font-family: Verdana;
}
.style3 {
				font-size: xx-small;
}
.style4 {
				background-color: #75ACFF;
}
.style5 {
				font-family: Verdana;
				font-size: xx-small;
				background-color: #75ACFF;
}
</style>
</head>
<% 
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetActivityContentData" & Session("ID")
cmd.Parameters.Refresh()
cmd("@ID") = Session("ActRetVal")
If Session("ActReadOnly") Then cmd("@ReadOnly") = "Y"
set rs = cmd.execute()

ClgCode = rs("ClgCode")

 %>
<div align="center">
				<center>
				<table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111" bgcolor="#9BC4FF">
								<form name="frmNotes" method="post" action="activity/actSubmit.asp">
								<input type="hidden" name="cmd" value="notes">
								<tr>
          <td width="100%" align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
          <table cellpadding="0" border="0">
						<tr>
										<td><img src="images/icon_activity_<% If Not IsNull(ClgCode) Then %>S<% Else %>O<% End If %>.gif"></td>
										<td><b><font face="Verdana" size="1"><%=getactivityContentLngStr("DtxtActivity")%>&nbsp;#<% If Not IsNull(ClgCode) Then Response.Write ClgCode Else Response.Write Session("ActRetVal") %>&nbsp;-&nbsp;<%=getactivityContentLngStr("LtxtContent")%></font></b></td>
						</tr>
			</table>
          </td>
								</tr>
								<tr>
												          <td width="100%">
												          <!--#include file="activityMenu.asp"--></td>
												        </tr>
								<tr>
												<td>
												<textarea <% If Session("ActReadOnly") Then %>readonly<% End If %> rows="17" name="Notes" cols="20" class="input" style="width: 100%"><%=myHTMLEncode(rs("Notes"))%></textarea></td>
								</tr>

								<!--#include file="activityBottom.asp"-->

								</form>
				</table>
				</center>
</div>

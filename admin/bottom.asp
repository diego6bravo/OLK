

&nbsp;</td>
	</tr>
	<tr>
		<td colspan="3" background="images/<%=Session("rtl")%>img_footer.jpg" height="49" valign="bottom" <% If Session("rtl") = "rtl/" Then %> style="background-position: top right;"<% End If %>>
		<font size="1" face="Verdana" color="#4AD1FF">OLK v.<%=OLKVerStr%>&nbsp;OBServer 
Emulator v.<%=R3VerStr%><% If myApp.VSystem <> "" Then %>&nbsp;(SBO <%=myApp.VSystem%>)<% End If %><br>
TopManage - Copyright 2002 - 2012</font></td>
		<td width="35%" background="images/admin_olk_new_r4_c4.jpg" valign="bottom">
&nbsp;</td>
	</tr>
</table>
<form name="frmChangeLng" method="post" action="" method="post">
<% For each itm in Request.Form
If itm <> "newLng" Then %>
<input type="hidden" name="<%=itm%>" value="<%=Server.HTMLEncode(Request.Form(itm))%>">
<% End If
Next
For Each itm in Request.QueryString
If itm <> "newLng" Then %>
<input type="hidden" name="<%=itm%>" value="<%=Server.HTMLEncode(Request.QueryString(itm))%>">
<% End If 
Next %>
<input type="hidden" name="newLng" value="">
</form>
<!--#include file="linkForm.asp"-->
<script type="text/javascript">doMsgBox();</script>
<div id="clearSpace"></div>
</body>

<% conn.close 
set rs = nothing %></html>
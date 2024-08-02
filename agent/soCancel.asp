<!--#include file="clientInc.asp"-->
<!--#include file="agentTop.asp"-->
<% If userType <> "V" Then Response.Redirect "default.asp" %>
<% 
If Not myApp.EnableOOPR Then Response.Redirect "unauthorized.asp"
addLngPathStr = "" %>
<!--#include file="lang/soCancel.asp" -->
<%
	If Session("SORetVal") <> "" Then
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKTransationCancel"
		cmd.Parameters.Refresh()
		cmd("@LogNum") = CLng(Session("SORetVal"))
		cmd.Execute()
       Session("SORetVal") = ""
    End If
                      %>
<div align="center">
  <center>
<table border="0" cellpadding="0" cellspacing="0" width="453">
  	<tr>
		<td>
		<p align="center">
		<img src="design/0/images/spacer.gif" width="17" height="1" border="0" alt=""></td>
		<td>
		<p align="center">
		<img src="design/0/images/spacer.gif" width="421" height="1" border="0" alt=""></td>
		<td>
		<p align="center">
		<img src="design/0/images/spacer.gif" width="15" height="1" border="0" alt=""></td>
		<td>
		<p align="center">
		<img src="design/0/images/spacer.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<% If Request("isUpdate") = "false" Then %>
	<tr>
		<td colspan="3">
		<p align="center">
		<img name="cancel_r1_c1" src="design/0/images/cancel_r1_c1.jpg" width="453" height="180" border="0" alt=""></td>
		<td>
		<p align="center">
		<img src="design/0/images/spacer.gif" width="1" height="180" border="0" alt=""></td>
	</tr>
	<% End If %>
	<tr>
		<td>
		<p align="center">
		&nbsp;</td>
		<td background="design/0/images/cancel_r2_c2.jpg">
		<p align="center"><% If Request("isUpdate") = "false" Then %><%=getsoCancelLngStr("LtxtSOCancel")%><% Else %><%=getsoCancelLngStr("LtxtSOUpdCancel")%><% End If %></td>
		<td>
		<p align="center">
		&nbsp;</td>
		<td>
		<p align="center">
		<img src="design/0/images/spacer.gif" width="1" height="36" border="0" alt=""></td>
	</tr>
	<tr>
		<td colspan="3">
		<p align="center">
		&nbsp;</td>
		<td>
		<p align="center">
		<img src="design/0/images/spacer.gif" width="1" height="47" border="0" alt=""></td>
	</tr>
</table>
  </center>
</div>
<!--#include file="agentBottom.asp"-->
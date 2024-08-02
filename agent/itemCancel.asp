<!--#include file="clientInc.asp"-->
<!--#include file="agentTop.asp"-->
<% If userType <> "V" Then Response.Redirect "default.asp" %>
<% If Not myApp.EnableOITM Then Response.Redirect "unauthorized.asp" %>
<% addLngPathStr = "" %>
<!--#include file="lang/itemCancel.asp" -->

<%	If Session("ItmRetVal") <> "" Then
		set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = connCommon
		cmd.CommandType = &H0004
		cmd.CommandText = "DBOLKTransationCancel"
		cmd.Parameters.Refresh()
		cmd("@LogNum") = CLng(Session("ItmRetVal"))
		cmd.Execute()
       Session("ItmRetVal") = ""
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
	<tr>
		<td colspan="3">
		<p align="center">
		<img name="cancel_r1_c1" src="design/0/images/cancel_r1_c1.jpg" width="453" height="180" border="0" alt=""></td>
		<td>
		<p align="center">
		<img src="design/0/images/spacer.gif" width="1" height="180" border="0" alt=""></td>
	</tr>
	<tr>
		<td>
		<p align="center">
		<img name="cancel_r2_c1" src="design/0/images/cancel_r2_c1.jpg" width="17" height="36" border="0" alt=""></td>
		<td background="design/0/images/cancel_r2_c2.jpg">
		<p align="center"><%=getitemCancelLngStr("LtxtItmCancel")%></td>
		<td>
		<p align="center">
		<img name="cancel_r2_c3" src="design/0/images/cancel_r2_c3.jpg" width="15" height="36" border="0" alt=""></td>
		<td>
		<p align="center">
		<img src="design/0/images/spacer.gif" width="1" height="36" border="0" alt=""></td>
	</tr>
	<tr>
		<td colspan="3">
		<p align="center">
		<img name="cancel_r3_c1" src="design/0/images/cancel_r3_c1.jpg" width="453" height="47" border="0" alt=""></td>
		<td>
		<p align="center">
		<img src="design/0/images/spacer.gif" width="1" height="47" border="0" alt=""></td>
	</tr>
</table>
  </center>
</div>
<!--#include file="agentBottom.asp"-->
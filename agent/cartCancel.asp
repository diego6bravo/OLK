<!--#include file="clientInc.asp"-->
<% Select Case userType
Case "C" %><!--#include file="clientTop.asp"-->
<% 
If (Session("UserName") = "-Anon-" or not optBasket) Then Response.Redirect "default.asp"
Case "V" %><!--#include file="agentTop.asp"-->
<% 
If Not comDocsMenu Then Response.Redirect "unauthorized.asp"
End Select %>
<% addLngPathStr = "" %>
<!--#include file="lang/cartCancel.asp" -->
<%
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKCancelLog"
cmd.Parameters.Refresh()
If Session("RetVal") <> "" Then
	cmd("@LogNum") = Session("RetVal")
	cmd.execute()
	Session("RetVal") = ""
	Session("PayRetVal") = ""
ElseIf Session("PayRetVal") <> "" Then
	cmd("@LogNum") = Session("PayRetVal")
	cmd.execute()
	Session("PayRetVal") = ""
End If
ObjCode = cmd("@Object")
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
		<td colspan="3" valign="bottom">
		<p align="center">
		<img border="0" src="images/RecycleBin.gif" width="128" height="128"></td>
		<td>
		<p align="center">
		<img src="design/0/images/spacer.gif" width="1" height="180" border="0" alt=""></td>
	</tr>
	<tr>
		<td>
		<p align="center">
		<img name="cancel_r2_c1" src="design/0/images/cancel_r2_c1.jpg" width="17" height="36" border="0" alt=""></td>
		<td background="design/0/images/cancel_r2_c2.jpg">
		<p align="center"><%
		Select Case ObjCode
			Case 17
				response.write txtOrdr & " " & getcartCancelLngStr("LtxtCanceled")
			Case 23
				response.write txtQuote & " " & getcartCancelLngStr("LtxtCanceled2")
			Case 13
				response.write txtInv & " " & getcartCancelLngStr("LtxtCanceled2")
			Case 15
				response.write txtOdln & " " & getcartCancelLngStr("LtxtCanceled2")
			Case 24
				Response.Write txtOrct & " " & getcartCancelLngStr("LtxtCanceled")
			Case 203
				response.write txtODPIReq & " " & getcartCancelLngStr("LtxtCanceled2")
			Case 204
				response.write txtODPIInv & " " & getcartCancelLngStr("LtxtCanceled2")
		End Select
		%></td>
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
<% Select Case userType
Case "C" %><!--#include file="clientBottom.asp"-->
<% Case "V" %><!--#include file="agentBottom.asp"-->
<% End Select %>
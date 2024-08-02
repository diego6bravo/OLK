<% addLngPathStr = "client/" %>
<!--#include file="lang/clientsubmitconfirm.asp" --><%
If Request("s") <> "E" Then
	sql = "select ObjectCode, Command from R3_ObsCommon..TLOG where LogNum = " & Session("ConfRetVal")
	set rs = conn.execute(sql)
Else
	sql = "select ErrMessage, Command from R3_ObsCommon..TLOG where LogNum = " & Session("CrdRetVal")
	set rs = conn.execute(sql)
End If
%><div align="center">
				<center>
				<table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
								<tr>
												<td bgcolor="#9BC4FF">
												<table border="0" cellpadding="0" bordercolor="#111111" width="100%" id="AutoNumber1">
																<tr>
																				<td width="100%" bgcolor="#75ACFF">
																				<p align="center"><b><font face="Verdana" size="1"><%=getclientsubmitconfirmLngStr("LtxtBPConf")%></font></b></p>
																				</td>
																</tr>
																<tr>
																				<td width="100%">&nbsp;</td>
																</tr>
																<% If Request("s") = "E" Then %>
																<tr>
																				<td width="100%" align="center"><img src="images/errorIcon.gif"> </td>
																</tr>
																<% End If %>
																<tr>
																				<td width="100%">
																				<table cellpadding="0" cellspacing="0" border="0" width="100%">
																								<tr>
																												<td><font size="1" face="Verdana"><b><% 
																								          Select Case Request("s")
																								          	Case "S"
																								          		Select Case rs("Command")
																								          			Case "A" 
																								          				Response.Write Replace(getclientsubmitconfirmLngStr("LtxtBPOk"), "{0}", rs("ObjectCode"))
																										          	Case "U"
																											          	 Response.Write Replace(getclientsubmitconfirmLngStr("LtxtBPUpdOk"), "{0}", rs("ObjectCode"))
																										        End Select
																									       Case "H"
																									     		Response.Write Replace(getclientsubmitconfirmLngStr("LtxtWaitForConf"), "{0}", Session("ConfRetVal"))
																									       Case "E"
																									     		Select Case rs("Command")
																									     			Case "A"
																									    	 			Response.Write getclientsubmitconfirmLngStr("LtxtErrAddBP")
																											        Case "U"
																											        	Response.Write getclientsubmitconfirmLngStr("LtxtErrUpdBP")	
																											    End Select %><br>
																												<%=rs("ErrMessage")%> <% End Select %></b></font></td>
																								</tr>
																								<% If Request("s") = "S" Then %>
																								<tr>
																												<td>&nbsp;</td>
																								</tr>
																								<tr>
																												<td align="center"><input type="button" name="btnOpenData" value="<%=getclientsubmitconfirmLngStr("LtxtOpenBP")%>" onclick="window.location.href='operaciones.asp?cmd=datos&amp;card=<%=CleanItem(rs("ObjectCode"))%>'"></td>
																								</tr>
																								<% End If %>
																				</table>
																				</td>
																</tr>
																<% If Request("s") = "E" Then %>
																<tr>
																				<td width="100%" align="center"><input type="button" name="btnRestore" value="<%=getclientsubmitconfirmLngStr("DtxtRestore")%>" onclick="javascript:window.location.href='operaciones.asp?cmd=newClient'"> -&nbsp; <input type="button" name="btnRetry" value="<%=getclientsubmitconfirmLngStr("DtxtRetry")%>" onclick="javascript:window.location.href='operaciones.asp?cmd=newClientSubmit&amp;retry=Y'"> </td>
																</tr>
																<% End If %>
												</table>
												</td>
								</tr>
				</table>
				</center></div>
<%
If Session("NotifyAdd") Then
	Session("NotifyAdd") = False
	sql = "EXEC OLKCommon..DBOLKObjAlert" & Session("ID") & " " & Session("ConfRetVal") & ", " & Session("branch") & ", 'V', '" & getMyLng & "'"
	conn.execute(sql)
End If
	
'Session("RetVal") = ""
%>
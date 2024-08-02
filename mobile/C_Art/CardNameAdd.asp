<% If Session("RetVal") <> "" Then 
set rCName = Server.CreateObject("ADODB.RecordSet")
sql = "select OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRD', 'CardName', CardCode, CardName) CardName from OCRD where CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "'"
set rCName = conn.execute(sql) %>
<tr><td align="right"><font face="Verdana" size="1"><b><%=rCName ("CardName")%></b></font></td></tr>
<% set rCName = nothing
End If %>
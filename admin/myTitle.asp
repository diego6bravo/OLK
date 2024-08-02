<% sql = "select T0.LineIndex, IsNull(T1.AlterName, T0.Name) Name, T0.Query, T0.Row, T0.Col, T0.ShowName " & _
      "from OLKDocAddHdr T0 " & _
      "left outer join OLKDocAddHdrAlterNames T1 on T1.LineIndex = T0.LineIndex and T1.LanID = " & Session("LanID") & " " & _
      "where T0.Access "
      
If Not Session("OLKAdmin") Then
	sql = sql & " in ('T','" & userType & "') "
Else
	sql = sql & " <> 'D' "
End If
      
sql = sql & "order by T0.Row, T0.Col"
      set rd = conn.execute(sql)
      If Not rd.eof Then
      
		sql = ""
		
		do while not rd.eof
			If sql <> "" Then sql = sql & ", "
			sql = sql & "(" & rd("Query") & ") 'Col" & rd("LineIndex") & "'"
		rd.movenext
		loop
		rd.movefirst
		
		sql = "select " & sql & " from OADM T0"
		set ra = Server.CreateObject("ADODB.RecordSet")
		set ra = conn.execute(sql) %>
	<tr>
		<td class="FirmTbl">
		<% lastRow = -1
		do while not rd.eof
		If CInt(rd("Row")) <> lastRow Then
			If lastRow <> -1 Then
				Response.Write "</tr></table>"
			End If
			lastRow = CInt(rd("Row"))
			Response.Write "<table cellpadding=""0"" cellspacing=""0"" border=""0""><tr>"
		End If %>
		<% If rd("ShowName") = "Y" Then %><td class="FirmTbl" style="padding-left: 2px; padding-right: 2px;"><b><%=rd("Name")%></b></td><% End If %>
		<td class="FirmTbl" style="padding-left: 2px; padding-right: 2px;"><%=ra("Col" & rd("LineIndex"))%></td>
		<% rd.movenext
		loop
		Response.Write "</tr></table>" %>
		</td>
	</tr>
	<% End If %>
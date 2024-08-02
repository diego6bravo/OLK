<%
set rx = Server.CreateObject("ADODB.RecordSet")
sql = "select T0.RowQuery, IsNull(T1.AlterRowName, T0.RowName) RowName, Align " & _
"from OLKCMREP T0 " & _
"left outer join OLKCMREPAlterNames T1 on T1.RowType = T0.RowType and T1.LineIndex = T0.LineIndex and T1.LanID = " & Session("LanID") & " " & _
"where T0.RowActive = 'Y' and T0.ShowV = 'Y' and RowQuery is not null and Main = 'Y'"
rx.open sql, conn, 3, 1
if not rx.eof then
	sql = "declare @LogNum int set @LogNum = " & Session("RetVal") & " " & _
	"declare @CardCode nvarchar(15) set @CardCode = N'" & saveHTMLDecode(Session("username"), False) & "' " & _
	"declare @LanID int set @LanID = " & Session("LanID") & " " & _
	"select "
	do while not rx.eof
		sql = sql & "(" & rx("RowQuery") & ") As N'" & Replace(rx("RowName"), "'", "''") & "'"
	rx.movenext
	loop
	sql = QueryFunctions(sql)
	set rx = conn.execute(sql)
	If Not rx.Eof Then %><table cellpadding="0" cellspacing="0" border="0" align="center" style="cursor: pointer;" onclick="window.location.href='operaciones.asp?cmd=cart_cp';"><tr><td><a href="operaciones.asp?cmd=cart_cp"><img border="0" src="images/<%=Session("rtl")%>flechaselec.gif"></a></td><td><font face="Verdana" size="2"><%=rx(0)%></font></td></tr></table><% End If 
End If %>
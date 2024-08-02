<%
Sub SaveAddress

	set conn=Server.CreateObject("ADODB.Connection")
	conn.Provider=olkSqlProv
	conn.Open  "Provider=SQLOLEDB;charset=utf8;" & _
	          "Data Source=" & olkip & ";" & _
	          "Initial Catalog=" & Session("olkdb") & ";" & _
	          "Uid=" & olklogin & ";" & _
	          "Pwd=" & olkpass & ""

	If Request("BillToCode") <> "" Then BillToCode = "N'" & Request("BillToCode") & "'" Else BillToCode = "NULL"
	If Request("ShipToCode") <> "" Then ShipToCode = "N'" & Request("ShipToCode") & "'" Else ShipToCode = "NULL"
	sql = "update R3_ObsCommon..TDOC set PayToCode = " & BillToCode & ", ShipToCode = " & ShipToCode & " where LogNum = " & Session("RetVal")
	conn.execute(sql)
	
End Sub
%>
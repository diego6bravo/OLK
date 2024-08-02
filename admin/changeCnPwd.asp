<% On Error Resume Next %>
<html>

<!--#include file="conn.asp" -->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>Conexión SQL</title>
</head>
<%
If Request.Form.Count > 0 Then
set conn = server.createobject("ADODB.Connection")
SqlStr = "Provider=SQLOLEDB;charset=utf8;" & _
          "Data Source=" & olkip & ";" & _
          "Initial Catalog=OLKCommon;" & _
          "Uid=" & Request("olklogin") & ";" & _
          "Pwd=" & Request("olkpassword") & ""
conn.open sqlStr
If Err.Number = 0 Then
conn.close
'ErrDesc = "Conexión establecida con exito"
CreateConn()
If Request("rAction") = "admin" Then
	response.redirect "default.asp"
ElseIf Request("rAction") = "c_p" Then
	response.redirect ""
ElseIf Request("rAction") = "pocket" Then
	response.redirect "mobile/"
End If
ElseIf Err.Number = -2147217843 Then
	ErrDesc = "Usuario o contraseña erronea, intente de nuevo"
Else
	ErrDesc = Err.Description
End If
End If

Public Sub CreateConn()
filePath = Server.MapPath("conn.asp")
dim fs, f
set fs = server.createobject("Scripting.FileSystemObject")
set f = fs.createtextfile(filePath,true)
f.writeline("<%")
f.writeline("olkip = """ & olkip & """")
f.writeline("olklogin = """ & Request("olklogin") & """")
f.writeline("olkpass = """ & Request("olkpassword") & """")
f.write("%")
f.write(">")
f.close
set f = nothing
set fs = nothing
End Sub
          %>
<body>

<div align="center">
	<table border="0" cellpadding="0" width="400" id="table1">
		<tr>
			<td>
			<p align="center"><b><font face="Verdana" size="2">Error de conexión 
			de SQL</font></b></td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">La conexión de SQL no pudo ser 
			establecida ya que la contraseña del usuario del SQL ha cambiado, 
			por favor introduzca el usuario y contraseña para conectarse con el 
			SQL. Si usted no tiene esta contraseña, comuníquese con su 
			administrador de sistemas o llame al centro de servicio de TopManage 
			OLK al 236-8812.</font></td>
		</tr>
		<form method="POST" action="changeCnPwd.asp">
		<tr>
			<td>
			<div align="center">
				<table border="0" cellpadding="0" width="100%" id="table2">
					<tr>
						<td width="100"><font face="Verdana" size="1">Usuario:</font></td>
						<td><font size="1" face="Verdana">
						<input type="text" name="olkLogin" size="20" style="font-family: Verdana; font-size: 10px; width: 100%"></font></td>
					</tr>
					<tr>
						<td width="100"><font face="Verdana" size="1">
						Contraseña:</font></td>
						<td><font size="1" face="Verdana">
						<input type="password" name="olkPassword" size="20" style="font-family: Verdana; font-size: 10px; width: 100%"></font></td>
					</tr>
					<% If ErrDesc <> "" Then %>
					<tr>
						<td width="100"><font face="Verdana" size="1">Mensaje de 
						Error:</font></td>
						<td>
						<font face="Verdana" size="1"><%=ErrDesc%></font></td>
					</tr>
					<% End If %>
					<tr>
						<td width="100">&nbsp;</td>
						<td>
						<p align="center"><font size="1" face="Verdana">
						<input type="submit" value="Guardar" name="B1" style="font-family: Verdana; font-size: 10px"></font></td>
					</tr>
				</table>
			</div>
			</td>
		</tr>
			<input type="hidden" name="rAction" value="<%=Request("rAction")%>">
		</form>
	</table>
</div>

</body>

</html>
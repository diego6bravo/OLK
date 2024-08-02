<%@Page Language="VB" Debug="true" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Xml" %>
<%@ Import Namespace="Microsoft.Win32" %>
<%@ Import Namespace="Microsoft.SqlServer.Management.Common" %>
<%@ Import Namespace="Microsoft.SqlServer.Management.Smo" %>
<!--#include file="lang/adminDocConfUpdate.aspx" -->
<script runat="server">
Dim olkip, olklogin, olkpass, olkSqlProv, BackupPath, dbName, licip, licport, userType as String
Dim dbID, obj as Integer
Dim Err as Boolean
Dim ErrMsg As String = ""

Function GetPosIndex(ByVal Value As String) As Integer
	Dim retVal As Integer
	
    retVal = Value.ToLower().IndexOf("select @error, @error_message")
    
    Return retVal

End Function

Sub StartUpdate()
	dbName = Request("dbName")
	obj = CInt(Request("obj"))
	dbID = CInt(Request("ID"))

	Dim sqlCn as new SqlConnection("Server=" & olkip & ";uid=" & olklogin & ";pwd=" & olkpass & ";Database=R3_ObsCommon")
	Dim sqlCnDB as new SqlConnection("Server=" & olkip & ";uid=" & olklogin & ";pwd=" & olkpass & ";Database=" & dbName)
	sqlCn.Open()
	sqlCnDB.Open()
	
	Dim sqlCm as SqlCommand
	
	Dim sqlTran as SqlTransaction = sqlCn.BeginTransaction()
	
	Try
    
	    Dim server As New Server(new ServerConnection(olkip, olklogin, olkpass))
	    Dim db As Database = server.Databases("R3_ObsCommon")
	    
        Dim proc As StoredProcedure = db.StoredProcedures("OBSSp_TransactionNotification")

        Dim olkPosStr As String = proc.Script()(2).Replace("CREATE PROCEDURE", "ALTER PROCEDURE")
        If olkPosStr.IndexOf("@AutoGenQry") = -1 Then
        	olkPosStr = olkPosStr.Insert(GetPosIndex(olkPosStr), VbNewLine & "/* { START OLK Variables } */ declare @AutoGenQry nvarchar(max) declare @AutoGenCode nvarchar(20) /* { END OLK Variables } */" & VbNewLine)
        End If

       	Dim sqlCmDB as New SqlCommand("select Active, GenQry from OLKGenCode where ObjCode = @ObjCode", sqlCnDB)
       	sqlCmDB.Parameters.Add("@ObjCode", SqlDbType.Int).Value = obj
       	
       	Dim sqlDr as SqlDataReader = sqlCmDB.ExecuteReader()
       	sqlDr.Read()
       	
       	Dim enableAutoGenCode as Boolean = sqlDr(0) = "Y"
       	Dim autoGenQry as String = sqlDr(1).ToString()
       	
       	sqlDr.Close()

		Dim GenType As String = ""
		Dim GenSize As Integer
		Dim R3Type As String
		Dim FldCode As String
		
		Select Case obj
			Case 2
				GenType = "OCRD"
				GenSize = 15
				R3Type = "TCRD"
				FldCode = "Card"
			Case 4
				GenType = "OITM"
				GenSize = 20
				R3Type = "TITM"
				FldCode = "Item"
		End Select
		
		Dim startIndex as Integer = olkPosStr.IndexOf("/* { START OLK " & dbName & " Auto Gen " & GenType & " } */")
		If startIndex <> -1 Then
			Dim endIndex as Integer = olkPosStr.IndexOf("/* { END OLK " & dbName & " Auto Gen " & GenType & " } */")
			olkPosStr = olkPosStr.Remove(startIndex, endIndex - startIndex + 32 + Len(dbName))
		End If
		
		If enableAutoGenCode Then
			Dim strGenQry as String = "/* { START OLK " & dbName & " Auto Gen " & GenType & " } */ " & VbNewLine & _  
									"If @dbName = '" & dbName & "' and @transaction_type = 'A' and @ObjectCode = '" & obj & "' begin " & VbNewLine & _  
									"	If (select AppID from TLOGControl where LogNum = @LogNum) = 'TM-OLK' begin " & VbNewLine & _  
									"		set @AutoGenQry = N'use [" & dbName & "] set @AutoGenCode = (" & autoGenQry.Replace("'", "''") & ")' " & VbNewLine & _  
									"		EXEC sp_executesql @AutoGenQry, N'@LogNum int, @AutoGenCode nvarchar(" & GenSize & ") out', @LogNum = @LogNum, @AutoGenCode = @AutoGenCode out " & VbNewLine & _  
									"		update R3_ObsCommon.." & R3Type & " set " & FldCode & "Code = @AutoGenCode where LogNum = @LogNum " & VbNewLine & _  
									"	End " & VbNewLine & _  
									"End  " & VbNewLine & _  
									"/* { END OLK " & dbName & " Auto Gen " & GenType & " } */ " 
									
        	olkPosStr = olkPosStr.Insert(GetPosIndex(olkPosStr), VbNewLine & strGenQry & VbNewLine)
		End If
		
       	sqlCm = New SqlCommand(olkPosStr, sqlCn, sqlTran)
       	sqlCm.ExecuteNonQuery()
    	
    	sqlTran.Commit()
    	
    Catch ex As Exception
        ErrMsg &= "Err upgrading database:" & VbNewLine & ex.Message
        Err = True
        sqlTran.Rollback()
        Exit Sub
    End Try
    
    sqlCn.Close()
	sqlCnDB.Close()
End Sub
</script>
<!--#include file="conn.asp"-->	
<% 
StartUpdate()
If Not Err Then
	Response.Redirect("adminDocConf.asp?object=" & obj)
Else %>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>New Page 1</title>
<style type="text/css">
.style1 {
	border: 1px solid #3580A8;
	background-color: #D7EFFD;
	font-family: Verdana;
	font-size: x-small;
	color: #3580A8;
}
</style>
</head>

<body style="text-align: center">

<table border="0" cellpadding="0" cellspacing="0" width="497">
<!-- fwtable fwsrc="Untitled" fwbase="ventana.gif" fwstyle="FrontPage" fwdocid = "1685812017" fwnested=""0" -->
  	<tr>
		<td>
		<p align="center">
		<img src="images/spacer.gif" width="1" height="1" border="0" alt=""></td>
		<td>
		<p align="center">
		<img src="images/spacer.gif" width="495" height="1" border="0" alt=""></td>
		<td>
		<p align="center">
		<img src="images/spacer.gif" width="1" height="1" border="0" alt=""></td>
		<td>
		<p align="center">
		<img src="images/spacer.gif" width="1" height="1" border="0" alt=""></td>
	</tr>
	<tr>
		<td colspan="3" class="style1">
		ADMIN DOC CONF UPDATE</td>
		<td>
		<p align="center">
		<img src="images/spacer.gif" width="1" height="15" border="0" alt=""></td>
	</tr>
	<tr>
		<td background="images/ventana_r2_c1.gif">
		<p align="center">
		<img name="ventana_r2_c1" src="images/ventana_r2_c1.gif" width="1" height="263" border="0" alt=""></td>
		<td bgcolor="#FFFFFF" background="images/ventana_r2_c2.gif" style="font-size: x-small; font-family: Verdana; color: #3580A8; padding: 10px;"><%=ErrMsg%><br/>
		<p align="center"><input type="button" name="btnGoBack" value="Return" onclick="javascript:window.location.href='adminDocConf.asp?Object=<%=obj%>';" style="color: #68A6C0; font-family: Tahoma; border: 1px solid #68A6C0; background-color: #E5F1FF; font-size:10px; height:23; font-weight:bold"></p></td>
		<td background="images/ventana_r2_c3.gif">
		<p align="center">
		<img name="ventana_r2_c3" src="images/ventana_r2_c3.gif" width="1" height="263" border="0" alt=""></td>
		<td>
		<p align="center">
		<img src="images/spacer.gif" width="1" height="263" border="0" alt=""></td>
	</tr>
</table>

</body>

</html>
<% End if %>
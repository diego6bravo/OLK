<%@Page Language="VB" Debug="true" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Xml" %>
<%@ Import Namespace="Microsoft.Win32" %>
<%@ Import Namespace="Microsoft.SqlServer.Management.Common" %>
<%@ Import Namespace="Microsoft.SqlServer.Management.Smo" %>
<!--|P:LangLink|-->
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
	dbID = CInt(Request("ID"))

	Dim sqlCn as new SqlConnection("Server=" & olkip & ";uid=" & olklogin & ";pwd=" & olkpass & ";Database=" & dbName)
	sqlCn.Open()
	
	Dim sqlCm as SqlCommand
	
	Dim sqlTran as SqlTransaction = sqlCn.BeginTransaction()
	
	Try
    
	    Dim server As New Server(new ServerConnection(olkip, olklogin, olkpass))
	    Dim db As Database = server.Databases(dbName)
	    
        Dim proc As StoredProcedure = db.StoredProcedures("SBO_SP_TransactionNotification")

        Dim olkPosStr As String = proc.Script()(2).Replace("CREATE proc", "ALTER proc")
        If olkPosStr.IndexOf("DBOLKCheckDraftControl") = -1 Then
        	olkPosStr = olkPosStr.Insert(GetPosIndex(olkPosStr), VbNewLine & "/* { START OLK Draft Control } */ If @object_type = '112' and @transaction_type in ('A', 'U') Begin EXEC OLKCommon.dbo.DBOLKCheckDraftControl" & dbID & " @Entry = @list_of_cols_val_tab_del End /* { END OLK Draft Controls } */" & VbNewLine)
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
End Sub
</script>
<!--#include file="conn.asp"-->	
<% 
StartUpdate()
If Not Err Then
	Response.Redirect("adminDocFlow.asp?FlowID=" & Request("FlowID"))
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
		ADMIN DOC FLOW  UPDATE</td>
		<td>
		<p align="center">
		<img src="images/spacer.gif" width="1" height="15" border="0" alt=""></td>
	</tr>
	<tr>
		<td background="images/ventana_r2_c1.gif">
		<p align="center">
		<img name="ventana_r2_c1" src="images/ventana_r2_c1.gif" width="1" height="263" border="0" alt=""></td>
		<td bgcolor="#FFFFFF" background="images/ventana_r2_c2.gif" style="font-size: x-small; font-family: Verdana; color: #3580A8; padding: 10px;"><%=ErrMsg%><br/>
		<p align="center"><input type="button" name="btnGoBack" value="Return" onclick="javascript:window.location.href='adminDocFlow.asp?FlowID=<%=Request("FlowID")%>';" style="color: #68A6C0; font-family: Tahoma; border: 1px solid #68A6C0; background-color: #E5F1FF; font-size:10px; height:23; font-weight:bold"></p></td>
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
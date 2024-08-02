<!--#include file="lang/UDFQry.asp" -->

<link rel="stylesheet" href="Reportes/style.css">
<% 
Select Case Request("returnCmd")
	Case "cartopt"
		TableID = "OINV"
	Case "cartEditLine"
		TableID = "INV1"
	Case "newClientUDF"
		TableID = "OCRD"
	Case "activityUDF"
		TableID = "OCLG"
	Case "newClientContact"
		TableID = "OCPR"
	Case "newClientAddress"
		TableID = "CRD1"
End Select

sql = "select IsNull(T1.AlterDescr, T0.Descr) Descr, T2.SqlQuery, T2.SqlQueryField " & _
"from CUFD T0 " & _
"left outer join OLKCUFDAlterNames T1 on T1.TableID = T0.TableID and T1.FieldID = T0.FieldID and T1.LanID = " & Session("LanID") & " " & _
"inner join OLKCUFD T2 on T2.TableID = T0.TableID and T2.FieldID = T0.FieldID " & _
"where T0.TableID = '" & TableID & "' and T0.AliasID = N'" & Request("editVar") & "'"
set rs = conn.execute(sql)
set rd = Server.CreateObject("ADODB.RecordSet")
fldDescr = rs("Descr")
fldQuery = rs("SqlQuery")
fldQueryField = rs("SqlQueryField")
rs.close

	LogNum = -1
	
	Select Case Request("returnCmd")
		Case "cartopt"
			Title = getUDFQryLngStr("LtxtShopCart")
			LogNum = Session("RetVal")
		Case "cartEditLine"
			Title = getUDFQryLngStr("LtxtShopCart")
			LogNum = Session("RetVal")
		Case "newClientUDF"
			Title = getUDFQryLngStr("DtxtBP")
			LogNum = Session("CrdRetVal")
		Case "activityUDF"
			Title = getUDFQryLngStr("DtxtActivity")
			LogNum = Session("ActRetVal")
		Case "newClientContact"
			Title = getUDFQryLngStr("DtxtContact")
			LogNum = Session("CrdRetVal")
		Case "newClientAddress"
			Title = getUDFQryLngStr("DtxtAddress")
			LogNum = Session("CrdRetVal")
	End Select
	
	sqlSmall = "declare @LanID int set @LanID = " & Session("LanID") & " declare @LogNum int set @LogNum = " & LogNum & " " & _
		"declare @dbName nvarchar(100) set @dbName = '" & Session("olkdb") & "' " & _
		"declare @branch int set @branch = " & Session("branch") & " declare @SlpCode int set @SlpCode = " & Session("vendid") & " " & _
		"declare @CardCode nvarchar(15) set @CardCode = N'" & Session("username") & "' "
	
	If Request("returnCmd") = "cartopt" or Request("returnCmd") = "cartEditLine" Then
		sqlSmall = sqlSmall & "declare @PriceList int set @PriceList = " & Session("PList") & " "
	End If
	
	If Request("returnCmd") = "cartEditLine" Then
		sqlSmall = sqlSmall & "declare @ItemCode nvarchar(15) set @ItemCode = (select ItemCode from R3_ObsCommon..DOC1 where LogNum = @LogNum and LineNum = " & Request("LineNum") & ") " & _
							"declare @WhsCode nvarchar(8) set @WhsCode = (select WhsCode from R3_ObsCommon..DOC1 where LogNum = @LogNum and LineNum = " & Request("LineNum") & ") "
	End If

	sqlSmall = sqlSmall & fldQuery

set rs = conn.execute(sqlSmall)
 %>

<script language="javascript">
function setVal(val)
{
	document.frmUFDCal.U_<%=Request("editVar")%>.value=val;
	document.frmUFDCal.cmd.value='<%=Request("returnCmd")%>';
	document.frmUFDCal.submit();
}
</script>
<table border="0" cellspacing="0" width="100%" id="table1">
	<tr class="TblTltMnu">
		<td colspan="2"><img border="0" src="images/arrow_menu.gif" width="9" height="6">&nbsp;<%=Title%> - <%=Descr%></td>
	</tr>
	<form method="POST" name="frmUFDCal" action="operaciones.asp">
	<tr class="TblAfueraMnu">
		<td colspan="2" align="center">
		<table border="0" cellpadding="0" width="100%" id="table2">
			<tr class="TblTltMnu">
			 <% For each Field in rs.Fields %>
				<td>
				<p align="center"><%=Server.HTMLEncode(Field.Name)%>&nbsp;</td>
			<% next %>
			</tr>
			<% do while not rs.eof
			fldValue = rs(CStr(fldQueryField)) %>
			<tr class="TblAfueraMnu">
			  <% For each Field in rs.Fields  
			  varx = varx + 1 %>
				<td width="175">
				<p><% If Not IsNull(Field) Then %><a href="javascript:setVal('<%=Server.HTMLEncode(fldValue)%>')"><%=Server.HTMLEncode(Field)%></a><% End If %>&nbsp;</td>
			<% next %>
			</tr>
		  <% varx = 0
		  rs.movenext
		  loop %>
		</table>
		</td>
	</tr>
	<tr class="TblAfueraMnu">
		<td colspan="2">
		<p align="center">
		</td>
	</tr>
		<% 	For each itm in Request.Form %>
		<input type="hidden" name="<%=itm%>" value="<%=Request(itm)%>">
		<% 	Next %>
	<tr class="TblAfueraMnu">
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	</tr>
	<tr class="TblAfueraMnu">
		<td>
		<p align="left">&nbsp;</td>
		<td>
		<p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
		<input type="submit" name="btnCancel" value="<%=getUDFQryLngStr("DtxtCancel")%>" onclick="javascript:document.frmUFDCal.cmd.value='<%=Request("returnCmd")%>';"></td>
	</tr>
	</form>
</table>
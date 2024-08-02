<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

set rs = server.createobject("ADODB.Recordset")

sql = 	"select 'drop ' + Case Type When 'P' Then 'Procedure' When 'F' Then 'Function' End + ' ' + 'DB' + TableID + '" & Session("ID") & "' " & _
		"from OLKTDB T0 " & _
		"where Deleted = 'N' and Type in ('P', 'F') and CUFDType is null and exists(select '' from sysobjects where name = 'DB' + T0.TableID + '" & Session("ID") & "' collate database_default) " & _
		"and TableID not in ('OLKPostObjectCreation', 'OLKSalesItemDetailsCustom', 'olkItemInvCustom', 'olkItemInvValCustom', 'OLKPostAddItemToDoc') and NewSAP in ('A','N') " & _
		"Group By TableID, Type "
set rs = connCommon.execute(sql)

do while not rs.eof
	sql = rs(0)
	connCommon.execute(sql)
rs.movenext
loop

sql = 	"declare @LawsSet nvarchar(2) set @LawsSet = (select LawsSet from CINF) " & _
		"If @LawsSet in ('CR', 'CL', 'GT', 'US', 'CA') begin set @LawsSet = 'MX' End " & _
		"Else If @LawsSet in ('AT', 'AU', 'BE', 'CH', 'CZ', 'DE', 'DK', 'ES', 'FI', 'FR', 'CN', 'CY', 'HU', 'IT', 'NL', 'NO', 'PL', 'PT', 'RU', 'SE', 'SK', 'GB', 'ZA') begin set @LawsSet = 'PA' End " & _
		"select Query, TableID, [Type]  " & _
		"from olkcommon..olktdb  " & _
		"where " & _
		"Upgrade = 'Y' and LawsSet in ('All', @LawsSet) and NewSAP in ('A','N') and " & _
		"Deleted = 'N' and Type in ('P', 'F') and CUFDType is null " & _
		"and TableID not in ('OLKPostObjectCreation', 'OLKSalesItemDetailsCustom', 'olkItemInvCustom', 'olkItemInvValCustom') " & _
		"order by [Type]"
set rs = conn.execute(sql)
do while not rs.eof
	sql = rs(0)
	sql = Replace(sql, "{dbID}", Session("ID"))
	sql = Replace(sql, "{dbName}", Session("olkdb"))
	connCommon.execute(sql)
rs.movenext
loop 

sql = "declare @MyTable table(ID nvarchar(100)) " & _  
"insert @MyTable select CUFD from OLKCommon..OLKTDB where CUFD is not null and CUFDType <> 'SQ' Group By CUFD " & _  
"declare @Table nvarchar(100) set @Table = (select Min(ID) from @MyTable) " & _  
"while @Table is not null begin " & _  
"	EXEC OLKCommon..DBOLKGetMyQuery" & Session("ID") & " @TableID = @Table " & _  
"	set @Table = (select Min(ID) from @MyTable where ID > @Table) " & _  
"End " & _  
"EXEC OLKCommon..DBOLKGenQry" & Session("ID") & " N'ItemRec' " & _  
"EXEC OLKCommon..DBOLKRestoreUAF" & Session("ID") & " " & _  
"EXEC OLKCommon..DBOLKGenSpecQry" & Session("ID") & " 'OLKPromotions' " 
conn.execute(sql)

rs.close
conn.close

Response.Redirect "adminSystem.asp?Restore=OK"
%>

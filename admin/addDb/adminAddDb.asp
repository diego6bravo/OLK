<!--#include file="../chkLogin.asp" -->
<!--#include file="../langIndex.inc" -->
<!--#include file="../conn.asp" -->

<% 
Server.ScriptTimeout = 100000

olkDB = Request.Form("dbName")
set conn=Server.CreateObject("ADODB.Connection")
conn.Provider=olkSqlProv
conn.Open  "Provider=SQLOLEDB;charset=utf8;" & _
          "Data Source=" & olkip & ";" & _
          "Initial Catalog=" & olkDB & ";" & _
          "Uid=" & olklogin & ";" & _
          "Pwd=" & olkpass & ";Application Name=OLK;"
          
set connCommon=Server.CreateObject("ADODB.Connection")
connCommon.Provider=olkSqlProv
connCommon.Open  "Provider=SQLOLEDB;charset=utf8;" & _
          "Data Source=" & olkip & ";" & _
          "Initial Catalog=OLKCommon;" & _
          "Uid=" & olklogin & ";" & _
          "Pwd=" & olkpass & ";Application Name=OLK;"

set rs = server.CreateObject("ADODB.RecordSet")

sql = "If (select Left(Version,2) from [" & olkDB & "]..CINF) <= 67 begin " & _
"	EXEC sp_executesql N' " & _
"	select 6 Version, LawsSet, 2 Language,  " & _
"	Case When Exists(select ''A'' from [" & olkDB & "]..cufd  " & _
"	where TableID = ''INV1'' and AliasID = ''LineMemo'') Then ''Y'' Else ''N'' End VerfyLineMemo  " & _
"	from [" & olkDB & "]..CINF' " & _
"end else begin " & _
"	EXEC sp_executesql N' " & _
"	select 6 Version, LawsSet, Case   " & _  
"		When Language = 1 Then 3 " & _  
"		When Language in (23,25) Then 2 " & _  
"		When Language in (19, 29) Then 6 " & _  
"		When Language = 22 Then 8 " & _  
"		When Language = 8 Then 4 " & _  
"		When Language = 10 Then 9 " & _  
"		When Language = 24 Then 13 " & _  
"		When Language = 26 Then 14 " & _  
"		When Language = 5 Then 15 " & _  
"		When Language = 27 Then 16 " & _  
"		Else 1 " & _  
"	End Language,  " & _
"	Case When Left(Version, 1) >= 8 or Exists(select ''A'' from [" & olkDB & "]..cufd  " & _
"	where TableID = ''INV1'' and AliasID = ''LineMemo'') Then ''Y'' Else ''N'' End VerfyLineMemo  " & _
"	from [" & olkDB & "]..CINF' " & _
"end "
	set rs = conn.execute(sql)
	If rs(0) >= 6 Then
		sqlVerStr = "N"
	ElseIf rs(0) = 5 THen
		sqlVerStr = "O"
	End If
	Session("SVer") = rs(0)
	LawsSet = rs("LawsSet")
	Language = rs("Language")
	arrLawsSet = "PA, MX, CL, CR, GT, IL, US, CA, AT, AU, BE, CH, CZ, DE, DK, ES, FI, FR, HU, IT, NL, NO, PL, PT, RU, SE, SK, GB, ZA, BR, CN, CY"
	If InStr(arrLawsSet, LawsSet) = 0 Then
		Response.Redirect "../admin.asp?cmd=home&dbName=" & Request("dbName") & "&dbErr=Y"
	End If
	If rs("VerfyLineMemo") = "N" Then 
		Response.Redirect "../admin.asp?cmd=home&LineMemoErr=Y&dbName=" & Request("dbName")
	End If
	Session("OlkDB") = olkDB

	Select Case LawsSet
		Case "CL", "CR", "GT", "US", "CA", "BR"
			LawsSet = "MX"
		Case "AT", "AU", "BE", "CH", "CZ", "DE", "DK", "ES", "FI", "FR", "CN", "CY", "HU", "IT", "NL", "NO", "PL", "PT", "RU", "SE", "SK", "GB", "ZA"
			LawsSet = "PA"
	End Select
ChkOldVersion()
CopyImg()
sql = "select Query, Type from olkcommon..olktdb where Deleted = 'N' and NewSAP in ('A','" & sqlVerStr & "') and LawsSet in ('All', '" & LawsSet & "') and CUFDType is null"
sqlAddStr = ""
rs.close
rs.open sql, conn, 3, 1

rs.Filter = "Type = 'T'"
do while not rs.eof
	sql = rs("Query")
	conn.execute(sql)
rs.movenext
loop

sql = "declare @Version nvarchar(10) set @Version = (select Version from olkcommon..olkAdminLogin) " & _
	  "declare @dbID int set @dbID = IsNull((select Max(ID)+1 from OLKCommon..OLKDBA), 0) " & _
	  "insert olkcommon..olkdba(ID, dbName, Active, Version) Values (@dbID, '" & olkDB & "', 'Y', @Version) " & _
	  "update olkcommon set Version = @Version"
conn.execute(sql)

sql = "select ID from OLKCommon..OLKDBA where dbName = '" & olkDB & "'"
set rd = Server.CreateObject("ADODB.RecordSet")
set rd = conn.execute(sql)
dbID = rd(0)

set connCreate = Server.CreateObject("ADODB.Connection")
connCreate.open "Provider=" & olkSqlProv & ";charset=utf8;" & _
          "Data Source=" & olkip & ";" & _
          "Initial Catalog=OLKCommon;" & _
          "Uid=" & olklogin & ";" & _
          "Pwd=" & olkpass & ""

rs.Filter = "Type = 'F'"
do while not rs.eof
	sql = Replace(Replace(rs("Query"), "{dbName}", olkDB), "{dbID}", dbID)
	connCreate.execute(sql)
rs.movenext
loop

rs.Filter = "Type = 'P'"
do while not rs.eof
	sql = Replace(Replace(rs("Query"), "{dbName}", olkDB), "{dbID}", dbID)
	connCreate.execute(sql)
rs.movenext
loop

rs.Filter = "Type = 'Q'"
do while not rs.eof
	sql = "declare @LanID int set @LanID = " & Language & " " & Replace(Replace(rs("Query"), "{dbName}", olkDB), "{dbID}", dbID)
	conn.execute(sql)
rs.movenext
loop
rs.Close

sql = "update OLKCommon set NatLng = (select Case Lower(LanSign) When 'es-la' then 'es' When 'en-us' Then 'en' when 'pt-br' then 'pt' else Lower(LanSign) end from OLKCommon..OLKLang where LanID = " & Language & ")"
conn.execute(sql)

conn.close
connCommon.close

response.redirect "../adminSubmit.asp?submitCmd=changeDb&dbName=" & dbID

Sub ChkOldVersion()
	sql = "select case when exists(select 'A' from [" & olkDB & "]..sysobjects where name = 'OLKCommon') Then 'Y' Else 'N' End Verfy"
	set rs = conn.execute(sql)
	If rs("Verfy") = "Y" Then
		sql = "select (select Version from OLKCommon..OLKAdminLogin) OLKv, (select Version from [" & olkDB & "]..OLKCommon) DBv"
		set rs = conn.execute(sql)
		If rs(0) = rs(1) Then
			Session("OlkDB") = olkDB
			sql = "declare @Version nvarchar(10) set @Version = (select Version from OLKCommon..olkAdminLogin) " & _
				  "declare @dbID int set @dbID = IsNull((select Max(ID)+1 from OLKCommon..OLKDBA), 0) " & _
				  "insert olkcommon..olkdba(ID, dbName, Active, Version) Values (@dbID, '" & olkDB & "', 'Y', @Version) " & _
				  "update olkcommon set Version = @Version"
			conn.execute(sql)
			response.redirect "../admin.asp?cmd=home"
		ElseIf rs(0) < rs(1) Then
			response.redirect "../admin.asp?cmd=home&dbVMErr=Y&dbName=" & Request("dbName")
		ElseIf rs(0) > rs(1) Then
			sql = "declare @Version nvarchar(10) set @Version = (select Version from [" & olkDB & "]..OLKCommon) " & _
				  "declare @dbID int set @dbID = IsNull((select Max(ID)+1 from OLKCommon..OLKDBA), 0) " & _
				  "insert olkcommon..olkdba(ID, dbName, Active, Version) Values (@dbID, '" & olkDB & "', 'Y', @Version) " & _
				  "update olkcommon set Version = @Version"
			conn.execute(sql)
			response.redirect "../updateDb/updateDb.asp?dbName=" & Request("dbName")
		End If
	End If
End Sub

Sub CopyImg()
	On Error Resume Next
	sql = "select BitmapPath from oadp"
	set rs = conn.execute(sql)
	If rs(0) <> "" Then
		filePath = rs(0)
		set fso = Server.CreateObject("Scripting.FileSystemObject")
		if fso.FolderExists(filePath) then
			Dim file(2)
			file(0) = "error.gif"
			file(1) = "n_a.gif"
			file(2) = "pcard.gif"
			On Error Resume Next
			For i = 0 to 2
			    fso.CopyFile Server.MapPath("NewImg/" & file(i)), filePath & "\" & file(i)
			next
		End If
	End If
End Sub

%>
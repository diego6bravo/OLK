<%@Page Language="VB" Debug="true" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Xml" %>
<%@ Import Namespace="Microsoft.Win32" %>
<%@ Import Namespace="Microsoft.SqlServer.Management.Common" %>
<%@ Import Namespace="Microsoft.SqlServer.Management.Smo" %>
<!--#include file="lang/update.aspx" -->
<script runat="server">
Dim olkip, olklogin, olkpass, olkSqlProv, BackupPath, dbName, licip, licport, userType as String
Dim dbID as Integer
Dim Err as Boolean
Dim ErrMsg As String = ""
Dim sqlCn as SqlConnection
Dim sqlCnCommon as SqlConnection
Dim sqlTran as SqlTransaction
Dim lastItem as String = ""

Dim oLic

Sub StartUpdate()
	dbName = Request("dbName")
	dbID = CInt(Request("dbID"))
	loadLanguage()

	sqlCn = new SqlConnection("Server=" & olkip & ";uid=" & olklogin & ";pwd=" & olkpass & _
	";Database=" & dbName)
	sqlCn.Open()

	sqlCnCommon = new SqlConnection("Server=" & olkip & ";uid=" & olklogin & ";pwd=" & olkpass & _
	";Database=OLKCommon")
	sqlCnCommon.Open()
	
	oLic = Server.CreateObject("TM.LicenceConnect.LicenceConnection")
	oLic.LicenceServer = licip
	oLic.LicencePort = licport
	
	Dim sqlStr as string = "if exists (select '' from olkcommon..sysobjects where name = 'OLKTempDB') " & _
							"drop table olkcommon..OLKTempDB"
	Dim sqlCm as new sqlCommand(sqlStr, sqlCn)
	sqlCm.ExecuteNonQuery()
		
    sqlStr = "CREATE TABLE [OLKCommon]..[OLKTempDB] ( " & _
                            "[OLDTableID] [nvarchar] (100) NOT NULL ," & _
                            "[NEWTableID] [nvarchar] (100) NOT NULL " & _
                            ") ON [PRIMARY]"
    dbName = Request("dbName")
    sqlCm = New SqlCommand(sqlStr, sqlCn)
    Try
        sqlCm.ExecuteNonQuery()
    Catch ex As Exception
        ErrMsg &= "Err creating OLKTempDB:" & VbNewLine & ex.Message
        Err = True
        Exit Sub
    End Try
    
    Try
        sqlCm = new sqlCommand("backup database [" & dbName & "] To DISK = 'olk_" & dbName & "" & now().Day & now().Month & now().Year & now().Hour & now().Minute & "' WITH FORMAT", sqlCn)
        sqlCm.CommandTimeout = 1800
        sqlCm.ExecuteNonQuery()
    Catch ex As Exception
        ErrMsg &= "Err creating database backup:" & VbNewLine & ex.Message
        Err = True
        Exit Sub
    End Try

	sqlTran = sqlCn.BeginTransaction()
	
	Try
    
        Dim i As Integer
        Dim dt As New DataTable()
        Dim sqlDa As New SqlDataAdapter("select name from sysobjects where name like 'TMPOLK%'", sqlCn)
        sqlDa.SelectCommand.Transaction = sqlTran
        sqlDa.Fill(dt)
        sqlDa.Dispose()
        Dim dv As New DataView(dt)
        For i = 0 To dv.Count - 1
            sqlCm = New SqlCommand("drop table [" & dv(i)("name") & "]", sqlCn, sqlTran)
            sqlCm.ExecuteNonQuery()
        Next

		ClearDB()
        UpdateDB()
    
    	sqlTran.Commit()
    	
    	
    Catch ex As Exception
        ErrMsg &= "Err upgrading database:" & VbNewLine & lastItem & ex.Message
        Err = True
        sqlTran.Rollback()
        
        Exit Sub
    End Try

    sqlCm = New SqlCommand("drop table [OLKCommon]..OLKTempDB ", sqlCn)
    sqlCm.ExecuteNonQuery()
    sqlCn.Close()
    sqlCnCommon.Close()
End Sub

Private Sub ClearDB()
    'Para eliminar las columnas U_OLK de las tablas OUSR y OPLN por problemas de compatibilidad con SAP 2004
    Dim sqlCm As SqlCommand
    Dim dt As New DataTable
    Dim sqlDa As New SqlDataAdapter("EXEC sp_helpconstraint 'OUSR', 'nomsg'", sqlCn)
    sqlDa.SelectCommand.Transaction = sqlTran
    sqlDa.Fill(dt)
    sqlDa.Dispose()
    Dim dv As New DataView(dt)
    dv.RowFilter = "constraint_name like '%OLK%'"
    Dim i As Integer
    For i = 0 To dv.Count - 1
        sqlCm = New SqlCommand("alter table OUSR drop constraint [" & dv(i)("constraint_name") & "]", sqlCn, sqlTran)
        sqlCm.ExecuteNonQuery()
    Next

    If New SqlCommand("select case when exists(select 'A' from sysobjects where Name = 'AUSR') then 'Y' Else 'N' End Verfy", sqlCn, sqlTran).ExecuteScalar() = "Y" Then
        dt = New DataTable
        sqlDa = New SqlDataAdapter("EXEC sp_helpconstraint 'AUSR', 'nomsg'", sqlCn)
        sqlDa.SelectCommand.Transaction = sqlTran
        sqlDa.Fill(dt)
        sqlDa.Dispose()
        dv = New DataView(dt)
        dv.RowFilter = "constraint_name like '%OLK%'"
        For i = 0 To dv.Count - 1
            sqlCm = New SqlCommand("alter table AUSR drop constraint [" & dv(i)("constraint_name") & "]", sqlCn, sqlTran)
            sqlCm.ExecuteNonQuery()
        Next
    End If

    dt = New DataTable
    sqlDa = New SqlDataAdapter("EXEC sp_helpconstraint 'OPLN', 'nomsg'", sqlCn)
    sqlDa.SelectCommand.Transaction = sqlTran
    sqlDa.Fill(dt)
    sqlDa.Dispose()
    dv = New DataView(dt)
    dv.RowFilter = "constraint_name like '%OLK%'"
    For i = 0 To dv.Count - 1
        sqlCm = New SqlCommand("alter table OPLN drop constraint [" & dv(i)("constraint_name") & "]", sqlCn, sqlTran)
        sqlCm.ExecuteNonQuery()
    Next

    dt = New DataTable
    sqlDa = New SqlDataAdapter("select name from sysobjects where name like 'TMPOLK%'", sqlCn)
    sqlDa.SelectCommand.Transaction = sqlTran
    sqlDa.Fill(dt)
    sqlDa.Dispose()
    dv = New DataView(dt)
    For i = 0 To dv.Count - 1
        sqlCm = New SqlCommand("drop table [" & dv(i)("name") & "]", sqlCn, sqlTran)
        sqlCm.ExecuteNonQuery()
    Next

    sqlCm = New SqlCommand("if exists(select 'A' from sysobjects where name = 'OLKCartOpt') and not exists(select 'A' from sysobjects where name = 'OLKCUFD') begin " & _
        "alter table OLKCartOpt add [TableID] nvarchar(20) default 'OINV' with values " & _
        "EXEC sp_rename 'OLKCartOpt', 'OLKCUFD' " & _
        "end ", sqlCn, sqlTran)
    sqlCm.ExecuteNonQuery()

    If New SqlCommand("select case when " & _
    "exists(select 'A' from syscolumns where id =  " & _
    "(select id from sysobjects where name = 'OLKDocConf') and name in ('CashAcct')) Then 'Y' Else 'N' End Verfy ", sqlCn, sqlTran).ExecuteScalar() = "N" Then

        sqlCm = New SqlCommand("alter table OLKDocConf add [CashAcct] [nvarchar] (20) NULL , [CheckAcct] [nvarchar] (20) NULL", sqlCn, sqlTran)
        sqlCm.ExecuteNonQuery()
        sqlCm = New SqlCommand("declare @CashAcct nvarchar(20) set @CashAcct = (select CashAcct from olkcommon) " & _
        "declare @CheckAcct nvarchar(20) set @CheckAcct = (select CheckAcct from olkcommon) " & _
        "update OLKDocConf set CashAcct = @CashAcct, CheckAcct = @CheckAcct ", sqlCn, sqlTran)
        sqlCm.ExecuteNonQuery()
    End If
    
    sqlCm = New SqlCommand("if exists(select 'A' from syscolumns where id = (select id from sysobjects where name = 'OLKMyCod') and name = 'Index') begin " & _
	"alter table OLKMyCod drop PK_OLKMyCod " & _
	"EXEC sp_rename 'OLKMyCod.Index', 'ID', 'COLUMN' " & _
	"alter table OLKMyCod add CONSTRAINT [PK_OLKMyCod] PRIMARY KEY CLUSTERED  " & _
	"	( " & _
	"		[Type] ASC, " & _
	"		[ID] ASC " & _
	"	) ON [PRIMARY] " & _
	"end ", sqlCn, sqlTran)
	sqlCm.ExecuteNonQuery()
	
	sqlCm = New SqlCommand("if exists(select 'A' from syscolumns where id = (select id from sysobjects where Name = 'OLKCommon') " & _
	"		and Name = 'AfterCartAdd') begin " & _
	"	alter table OLKCommon Add AfterCartAddC char(1) not null default 'N' with values " & _
	"	alter table OLKCommon Add AfterCartAddV char(1) not null default 'N' with values " & _
	"	EXEC sp_executesql N'update OLKCommon set AfterCartAddC = AfterCartAdd, AfterCartAddV = AfterCartAdd' " & _
	"	alter table OLKCommon drop constraint [DF_OLKCommon_AfterCartAdd] " & _
	"	alter table OLKCommon drop column AfterCartAdd " & _
	"end ", sqlCn, sqlTran)
	sqlCm.ExecuteNonQuery()
	
    sqlCm = New SqlCommand("if not exists(select 'A' from syscolumns where id = (select id from sysobjects where name = 'OLKCCart') and name = 'ColOrdr') begin " & _
    "	alter table OLKCCart add [ColOrdr] int not null default 1 with values " & _
    "	EXEC sp_executesql N'update OLKCCart set ColOrdr = LineIndex' " & _
    "end ", sqlCn, sqlTran)
    sqlCm.ExecuteNonQuery()
    
    sqlCm = New SqlCommand("if not exists(select 'A' from syscolumns where id = (select id from sysobjects where name = 'OLKTCart') and name = 'ColOrdr') begin " & _
    "	alter table OLKTCart drop constraint [DF_OLKCPT_ColType] " & _
    "	alter table OLKTCart drop constraint [DF_OLKTCart_ColTypeRnd] " & _
    "	alter table OLKTCart drop constraint [PK_OLKTCart] " & _
    "	EXEC sp_rename [OLKTCart], [tmpOLKTCart] " & _
    "	alter table tmpOLKTCart add [newIndex] [int] IDENTITY (1, 1) NOT NULL " & _
    "	CREATE TABLE [OLKTCart] ( " & _
    "		[LineIndex] [int] NOT NULL , " & _
    "		[ColName] [nvarchar] (50) NOT NULL , " & _
    "		[ColQuery] [ntext] NOT NULL , " & _
    "		[ColCustom] [char] (1) NOT NULL , " & _
    "		[ColType] [char] (1) NOT NULL , " & _
    "		[ColTypeRnd] [char] (1) NOT NULL , " & _
    "		[ColAccess] [char] (1) NOT NULL , " & _
    "		[ColAlign] [nvarchar] (50) NOT NULL , " & _
    "		[ColOrdr] [int] NOT NULL , " & _
    "		[ColIndex] [char] (1) NOT NULL , " & _
    "		CONSTRAINT [PK_OLKTCart] PRIMARY KEY  CLUSTERED  " & _
    "		( " & _
    "			[LineIndex] " & _
    "		)  ON [PRIMARY]  " & _
    "	) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY] " & _
    "	EXEC sp_executesql N' " & _
    "	insert OLKTCart(LineIndex, ColName, ColQuery, ColCustom, ColType, ColTypeRnd, ColAccess, ColAlign, ColOrdr, ColIndex) " & _
    "	select newIndex, ColName, ColQuery, ColCustom, ColType, ColTypeRnd, ColAccess, ColAlign, LineIndex, ColIndex " & _
    "	from tmpOLKTCart' " & _
    "	drop table [tmpOLKTCart] " & _
    "end ", sqlCn, sqlTran)
    sqlCm.ExecuteNonQuery()
    
    sqlCm = New SqlCommand("if not exists(select 'A' from syscolumns where id = (select id from sysobjects where name = 'OLKNews') and name = 'newsSmallText') begin " & _
	"	if exists(select 'A' from syscolumns where id = (select id from sysobjects where name = 'OLKNews') and name = 'Status') begin " & _
	"	 	alter table OLKNews drop constraint DF_OLKNews_Status " & _
	"	end else begin " & _
	"		alter table OLKNews add [Status] [char] (1) NULL " & _
	"	end " & _
	" 	alter table OLKNews drop constraint PK_OLKNews " & _
	" 	EXEC sp_rename [OLKNews], [tmpOLKNews] 	CREATE TABLE [OLKNews] ( " & _
	" 		[newsIndex] [int] NOT NULL , " & _
	" 		[newsDate] [datetime] NOT NULL , " & _
	" 		[newsTitle] [nvarchar] (254) NOT NULL , " & _
	" 		[newsSmallText] [nvarchar] (155) NOT NULL ,  " & _
	"		[newsText] [ntext] NOT NULL , " & _
	"		[newsImg] [nvarchar] (200) NULL , " & _
	" 		[newsSource] [nvarchar] (254) NULL , " & _
	" 		[Status] [char] (1) NOT NULL CONSTRAINT [DF_OLKNews_Status] DEFAULT ('A'),  " & _
	"		CONSTRAINT [PK_OLKNews] PRIMARY KEY  CLUSTERED " & _
	"  		( " & _
	" 			[newsIndex] " & _
	" 		) WITH  FILLFACTOR = 90  ON [PRIMARY] " & _
	"  	) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY] " & _
	" 	EXEC sp_executesql N'insert OLKNews(newsIndex, newsDate, newsTitle, newsSmallText, newsText, newsImg, newsSource, Status) " & _
	"  			select newsIndex, newsDate, newsTitle, Convert(nvarchar(155),newsText), newsText, newsImg, newsSource, Status  " & _
	"			from tmpOLKNews " & _
	" 			drop table tmpOLKNews' " & _
	"end  ", sqlCn, sqlTran)
    sqlCm.ExecuteNonQuery()
    
    sqlCm = New SqlCommand("if exists(select '' from syscolumns where name = 'ColCustom' and id = (select id from sysobjects where name = 'OLKTCart')) begin " & _
	"	EXEC sp_executesql N'update OLKTCart set ColQuery = ''OITM.'' + Convert(nvarchar(100),ColQuery) where ColCustom = ''N''' " & _
	"end " & _
	"if exists(select '' from syscolumns where name = 'ColCustom' and id = (select id from sysobjects where name = 'OLKCCart')) begin " & _
	"	EXEC sp_executesql N'update OLKCCart set ColQuery = ''OITM.'' + Convert(nvarchar(100),ColQuery) where ColCustom = ''N''' " & _
	"end ", sqlCn, sqlTran)
    sqlCm.ExecuteNonQuery()
    
    sqlCm = New SqlCommand("If Not Exists(select '' from syscolumns where id = (select id from sysobjects where name = 'OLKCommon')  " & _
	"	and name = 'showCxcOpenInvByC') begin  " & _
	"	Alter Table OLKCommon add  " & _
	"			showCxcOpenInvC char(1) default 'N' not null, " & _
	"			showCxcOpenInvByC nvarchar(50) default 'DocDate' not null, " & _
	"			showCxcIncTransC char(1) default 'N' not null, " & _
	"			showCxcDueDateC char(1) default 'N' not null " & _
	"	EXEC sp_executesql N'update OLKCommon set showCxcOpenInvC = showCxcOpenInv,  " & _
	"	showCxcOpenInvByC = showCxcOpenInvBy, showCxcIncTransC = showCxcIncTrans,  " & _
	"	showCxcDueDateC = showCxcDueDate'  " & _
	"End  ", sqlCn, sqlTran)
	sqlCm.ExecuteNonQuery()

    sqlCm = New SqlCommand("if not exists(select '' from syscolumns where id = (select id from sysobjects where name = 'OLKCUFD') and name = 'Order') begin " & _
    "	alter table OLKCUFD add [Order] [int] not null default 0 " & _
    "	declare @TableID nvarchar(20) set @TableID = (select Min(TableID) from OLKCUFD) " & _
    "	while @TableID is not null begin " & _
    "		declare @Pos char(1) set @Pos = (select Min(Pos) from OLKCUFD where TableID = @TableID) " & _
    "		declare @Order int " & _
    "		while @Pos is not null begin  " & _
    "			set @Order = 0 " & _
    "			declare @FieldID int set @FieldID = (select Min(FieldID) from OLKCUFD where TableID = @TableID and Pos = @Pos) " & _
    "			while @FieldID is not null begin " & _
    "				set @Order = @Order + 1 " & _
    "				declare @Qry nvarchar(4000) set @Qry = N'update OLKCUFD set [Order] = @Order where TableID = @TableID and FieldID = @FieldID' " & _
    "				EXEC sp_executesql @Qry, N'@Order int, @TableID nvarchar(20), @FieldID int', @Order = @Order, @TableID = @TableiD, @FieldID = @FieldID " & _
    "				set @FieldID = (select Min(FieldID) from OLKCUFD where TableID = @TableID and Pos = @Pos and FieldID > @FieldID) " & _
    "			End " & _
    "			set @Pos = (select Min(Pos) from OLKCUFD where TableID = @TableID and Pos > @Pos) " & _
    "		End " & _
    "		set @TableID = (select Min(TableID) from OLKCUFD where TableID > @TableID) " & _
    "	End " & _
    "End ", sqlCn, sqlTran)
    sqlCm.ExecuteNonQuery()
    
    sqlCm = New SqlCommand("if not exists(select 'A' from syscolumns where id = (select id from sysobjects where name = 'OLKCommon') and name = 'PrintPriceBefDiscount') " & _  
	"and exists(select 'A' from syscolumns where id = (select id from sysobjects where name = 'OLKCommon') and name = 'ShowLineDiscount') begin  " & _  
	"	ALTER TABLE OLKCommon add [PrintPriceBefDiscount] [char](1) NOT NULL DEFAULT 'N' WITH VALUES  " & _  
	"	ALTER TABLE OLKCommon add [PrintLineDiscount]	[char](1) NOT NULL DEFAULT 'Y' WITH VALUES  " & _  
	"	EXEC sp_executesql N'update OLKCommon set PrintPriceBefDiscount = ShowPriceBefDiscount, PrintLineDiscount = ShowLineDiscount'  " & _  
	"End  " , sqlCn, sqlTran)
    sqlCm.ExecuteNonQuery()
    
    sqlCm = New SqlCommand("declare @Version nvarchar(15) set @Version = (select Version from OLKCommon) " & _
    "if @Version < '1.91.17' begin " & _
    "	declare @LanID int set @LanID = (select LanID from OLKCommon..OLKLang where Left(LanSign, 2) = (select NatLng collate database_default from OLKCommon)) " & _
    "	declare @ID int set @ID = IsNull((select Max(ID) from OLKUAFControl), 0) " & _
    "	insert OLKUAFControl(ID, UserType, ExecAt, ObjectCode, ObjectEntry, Series, RequestDate, RequestUserSign, RequestLanID, RequestBranchID, ConfirmDate, ConfirmUserSign, Note, LogNum, Status) " & _
    "	select ROW_NUMBER() OVER(order by T0.LogNum)+@ID, 'V', 'D3', null, T0.LogNum, null, ConfRequestDate, UserSign, @LanID LanID, IsNull(ConfBranch, -1), null, null, null, null, 'O' " & _
    "	from R3_ObsCommon..TLOG T0 " & _
    "	inner join R3_ObsCommon..TLOGControl T1 on T1.LogNum = T0.LogNum " & _
    "	where T0.Company = db_name() and T0.Status = 'H' and T0.Object in (13,15,17,23) " & _
    "	and not exists(select '' from OLKUAFControl where ExecAt = 'D3' and ObjectEntry = T0.LogNum) " & _
    "	insert OLKUAFControl1(ID, FlowID, Note) " & _
    "	select T0.ID, T1.FlowID, T1.Note " & _
    "	from OLKUAFControl T0 " & _
    "	inner join OLKUAF3 T1 on T1.LogNum = T0.ObjectEntry " & _
    "	where T0.ExecAt = 'D3' " & _
    "	and not exists(select '' from OLKUAFControl1 where ID = T0.ID and FlowID = T1.FlowID) " & _
    "	set @ID = IsNull((select Max(ID) from OLKUAFControl), 0) " & _
    "	insert OLKUAFControl(ID, UserType, ExecAt, ObjectCode, ObjectEntry, Series, RequestDate, RequestUserSign, RequestLanID, RequestBranchID, ConfirmDate, ConfirmUserSign, Note, LogNum, Status) " & _
    "	select ROW_NUMBER() OVER(order by T0.LogNum)+@ID, 'V', 'R2', null, T0.LogNum, null, ConfRequestDate, UserSign, @LanID LanID, IsNull(ConfBranch, -1), null, null, null, null, 'O' " & _
    "	from R3_ObsCommon..TLOG T0 " & _
    "	inner join R3_ObsCommon..TLOGControl T1 on T1.LogNum = T0.LogNum " & _
    "	where T0.Company = db_name() and T0.Status = 'H' and T0.Object in (24) " & _
    "	and not exists(select '' from OLKUAFControl where ExecAt = 'R2' and ObjectEntry = T0.LogNum) " & _
    "	insert OLKUAFControl1(ID, FlowID, Note) " & _
    "	select T0.ID, T1.FlowID, T1.Note " & _
    "	from OLKUAFControl T0 " & _
    "	inner join OLKUAF3 T1 on T1.LogNum = T0.ObjectEntry " & _
    "	where T0.ExecAt = 'R2' " & _
    "	and not exists(select '' from OLKUAFControl1 where ID = T0.ID and FlowID = T1.FlowID) " & _
    "	set @ID = IsNull((select Max(ID) from OLKUAFControl), 0) " & _
    "	insert OLKUAFControl(ID, UserType, ExecAt, ObjectCode, ObjectEntry, Series, RequestDate, RequestUserSign, RequestLanID, RequestBranchID, ConfirmDate, ConfirmUserSign, Note, LogNum, Status) " & _
    "	select ROW_NUMBER() OVER(order by T0.LogNum)+@ID, 'V', 'A1', null, T0.LogNum, null, ConfRequestDate, UserSign, @LanID LanID, IsNull(ConfBranch, -1), null, null, null, null, 'O' " & _
    "	from R3_ObsCommon..TLOG T0 " & _
    "	inner join R3_ObsCommon..TLOGControl T1 on T1.LogNum = T0.LogNum " & _
    "	where T0.Company = db_name() and T0.Status = 'H' and T0.Object in (4) " & _
    "	and not exists(select '' from OLKUAFControl where ExecAt = 'A1' and ObjectEntry = T0.LogNum) " & _
    "	insert OLKUAFControl1(ID, FlowID, Note) " & _
    "	select T0.ID, T1.FlowID, T1.Note " & _
    "	from OLKUAFControl T0 " & _
    "	inner join OLKUAF3 T1 on T1.LogNum = T0.ObjectEntry " & _
    "	where T0.ExecAt = 'A1' " & _
    "	and not exists(select '' from OLKUAFControl1 where ID = T0.ID and FlowID = T1.FlowID) " & _
    "	set @ID = IsNull((select Max(ID) from OLKUAFControl), 0) " & _
    "	insert OLKUAFControl(ID, UserType, ExecAt, ObjectCode, ObjectEntry, Series, RequestDate, RequestUserSign, RequestLanID, RequestBranchID, ConfirmDate, ConfirmUserSign, Note, LogNum, Status) " & _
    "	select ROW_NUMBER() OVER(order by T0.LogNum)+@ID, 'V', 'C1', null, T0.LogNum, null, ConfRequestDate, UserSign, @LanID LanID, IsNull(ConfBranch, -1), null, null, null, null, 'O' " & _
    "	from R3_ObsCommon..TLOG T0 " & _
    "	inner join R3_ObsCommon..TLOGControl T1 on T1.LogNum = T0.LogNum " & _
    "	where T0.Company = db_name() and T0.Status = 'H' and T0.Object in (2) " & _
    "	and not exists(select '' from OLKUAFControl where ExecAt = 'C1' and ObjectEntry = T0.LogNum) " & _
    "	insert OLKUAFControl1(ID, FlowID, Note) " & _
    "	select T0.ID, T1.FlowID, T1.Note " & _
    "	from OLKUAFControl T0 " & _
    "	inner join OLKUAF3 T1 on T1.LogNum = T0.ObjectEntry " & _
    "	where T0.ExecAt = 'C1' " & _
    "	and not exists(select '' from OLKUAFControl1 where ID = T0.ID and FlowID = T1.FlowID) " & _
    "End ", sqlCn, sqlTran)
    Try
        sqlCm.ExecuteNonQuery()
    Catch ex As Exception

    End Try
    

    dt = New DataTable
    sqlDa = New SqlDataAdapter("declare @Version nvarchar(15) set @Version = (select Version from OLKCommon) " & _
                        "select SlpCode, [Authorization] from OLKAgentsAccess where @Version < '1.90.45' and [Authorization] is not null ", sqlCn)
    sqlDa.SelectCommand.Transaction = sqlTran
    sqlDa.Fill(dt)
    sqlDa.Dispose()
    sqlCm = New SqlCommand("update OLKAgentsAccess set [Authorization] = @Authorization where SlpCode = @SlpCode", sqlCn, sqlTran)
    sqlCm.Parameters.Add("@SlpCode", SqlDbType.Int)
    sqlCm.Parameters.Add("@Authorization", SqlDbType.VarChar, 8000)
    For i = 0 To dt.Rows.Count - 1
        Dim SlpCode As Integer = dt.Rows(i)("SlpCode")
        Dim strAut As String = dt.Rows(i)("Authorization")

        Dim sIndex As Integer = strAut.IndexOf("|P")
        
        If sIndex <> -1 Then

            strAut = strAut.Substring(0, sIndex) & "|S96%{Y}|S97%{Y}" & strAut.Substring(sIndex)

            sqlCm.Parameters("@SlpCode").Value = SlpCode
            sqlCm.Parameters("@Authorization").Value = strAut
            sqlCm.ExecuteNonQuery()
		
		End If
    Next
 

    dt = New DataTable
    sqlDa = New SqlDataAdapter("declare @Version nvarchar(15) set @Version = (select Version from OLKCommon) " & _
                        "select SlpCode, [Authorization] from OLKAgentsAccess where @Version < '1.91.49' and [Authorization] is not null ", sqlCn)
    sqlDa.SelectCommand.Transaction = sqlTran
    sqlDa.Fill(dt)
    sqlDa.Dispose()
    sqlCm = New SqlCommand("update OLKAgentsAccess set [Authorization] = @Authorization where SlpCode = @SlpCode", sqlCn, sqlTran)
    sqlCm.Parameters.Add("@SlpCode", SqlDbType.Int)
    sqlCm.Parameters.Add("@Authorization", SqlDbType.VarChar, 8000)
    For i = 0 To dt.Rows.Count - 1
        Dim SlpCode As Integer = dt.Rows(i)("SlpCode")
        Dim strAut As String = dt.Rows(i)("Authorization")

        Dim sIndex As Integer = strAut.IndexOf("|P")
        
        If sIndex <> -1 Then

            strAut = strAut.Substring(0, sIndex) & "|S175%{Y}|S174%{Y}" & strAut.Substring(sIndex)

            sqlCm.Parameters("@SlpCode").Value = SlpCode
            sqlCm.Parameters("@Authorization").Value = strAut
            sqlCm.ExecuteNonQuery()
		
		End If
    Next   
    
    dt = New DataTable
    sqlDa = New SqlDataAdapter("declare @Version nvarchar(15) set @Version = (select Version from OLKCommon) " & _
                        "select SlpCode, [Authorization] from OLKAgentsAccess where @Version < '1.90.58' and [Authorization] is not null ", sqlCn)
    sqlDa.SelectCommand.Transaction = sqlTran
    sqlDa.Fill(dt)
    sqlDa.Dispose()
    sqlCm = New SqlCommand("update OLKAgentsAccess set [Authorization] = @Authorization where SlpCode = @SlpCode", sqlCn, sqlTran)
    sqlCm.Parameters.Add("@SlpCode", SqlDbType.Int)
    sqlCm.Parameters.Add("@Authorization", SqlDbType.VarChar, 8000)
    For i = 0 To dt.Rows.Count - 1
        Dim SlpCode As Integer = dt.Rows(i)("SlpCode")
        Dim strAut As String = dt.Rows(i)("Authorization")

        Dim sIndex As Integer = strAut.IndexOf("|P")
        
        If sIndex <> -1 Then

            strAut = strAut.Substring(0, sIndex) & "|S99%{Y}|S100%{Y}|S101%{Y}|S102%{Y}|S103%{Y}" & strAut.Substring(sIndex)

            sqlCm.Parameters("@SlpCode").Value = SlpCode
            sqlCm.Parameters("@Authorization").Value = strAut
            sqlCm.ExecuteNonQuery()
		End If
    Next
    

    dt = New DataTable
    sqlDa = New SqlDataAdapter("declare @Version nvarchar(15) set @Version = (select Version from OLKCommon) " & _
                        "select SlpCode, [Authorization] from OLKAgentsAccess where @Version < '1.91.13' and [Authorization] is not null ", sqlCn)
    sqlDa.SelectCommand.Transaction = sqlTran
    sqlDa.Fill(dt)
    sqlDa.Dispose()
    If dt.Rows.Count > 0 Then
        sqlCm = New SqlCommand("update OLKAgentsAccess set [Authorization] = @Authorization where SlpCode = @SlpCode", sqlCn, sqlTran)
        sqlCm.Parameters.Add("@SlpCode", SqlDbType.Int)
        sqlCm.Parameters.Add("@Authorization", SqlDbType.VarChar, 8000)

        Dim arrConfAut As Object() = New Object() {New String() {"S77", "S114"}, New String() {"S45", "S116"}, New String() {"S78", "S118"}, New String() {"S44", "S121"}, _
                                        New String() {"S27", "S126"}, New String() {"S30", "S137"}, New String() {"S31", "S140"}, New String() {"S29", "S143"}, _
                                        New String() {"S32", "S145"}, New String() {"S34", "S147"}, New String() {"S35", "S148"}}

        For i = 0 To dt.Rows.Count - 1
            Dim SlpCode As Integer = dt.Rows(i)("SlpCode")
            Dim strAut As String = dt.Rows(i)("Authorization")

			Try
                Dim a As Integer
                For a = 0 To arrConfAut.Length - 1
                    Dim intPos As Integer

                    intPos = strAut.IndexOf(arrConfAut(a)(0))

                    Dim extract As String = strAut.Substring(intPos)

                    extract = extract.Substring(0, extract.IndexOf("|"))

                    Dim finalAut As String = extract

                    If finalAut.IndexOf("{A}") <> -1 Then
                        finalAut = finalAut.Remove(extract.IndexOf("{A}"), 3)
                        
                        finalAut = finalAut & "|" & arrConfAut(a)(1) & "%{Y}"
                    End If

                    strAut = strAut.Replace(extract, finalAut)

                Next

                sqlCm.Parameters("@SlpCode").Value = SlpCode
                sqlCm.Parameters("@Authorization").Value = strAut
                sqlCm.ExecuteNonQuery()
			Catch
			End Try

        Next

    End If
    

    dt = New DataTable
    sqlDa = New SqlDataAdapter("declare @Version nvarchar(15) set @Version = (select Version from OLKCommon) " & _
                        "select SlpCode, [Password] from OLKAgentsAccess where @Version < '1.92.72' and [Password] is not null ", sqlCn)
    sqlDa.SelectCommand.Transaction = sqlTran
    sqlDa.Fill(dt)
    sqlDa.Dispose()
    If dt.Rows.Count > 0 Then

        sqlCm = New SqlCommand("update OLKAgentsAccess set [Password] = @Password where SlpCode = @SlpCode", sqlCn, sqlTran)
        sqlCm.Parameters.Add("@SlpCode", SqlDbType.Int)
        sqlCm.Parameters.Add("@Password", SqlDbType.VarChar, 100)

        For i = 0 To dt.Rows.Count - 1
            Dim SlpCode As Integer = dt.Rows(i)("SlpCode")
            Dim strPwd As String = dt.Rows(i)("Password")

            sqlCm.Parameters("@SlpCode").Value = SlpCode
            sqlCm.Parameters("@Password").Value = oLic.GetEncPwd(strPwd)
            sqlCm.ExecuteNonQuery()

        Next

    End If
    
    dt = New DataTable
    sqlDa = New SqlDataAdapter("declare @Version nvarchar(15) set @Version = (select Version from OLKCommon) " & _
                        "select CardCode, [Password] from OLKClientsAccess where @Version < '1.92.72' and [Password] is not null ", sqlCn)
    sqlDa.SelectCommand.Transaction = sqlTran
    sqlDa.Fill(dt)
    sqlDa.Dispose()
    If dt.Rows.Count > 0 Then

        sqlCm = New SqlCommand("update OLKClientsAccess set [Password] = @Password where CardCode = @CardCode ", sqlCn, sqlTran)
        sqlCm.Parameters.Add("@CardCode ", SqlDbType.NVarChar, 15)
        sqlCm.Parameters.Add("@Password", SqlDbType.NVarChar, 100)

        For i = 0 To dt.Rows.Count - 1
            Dim cardCode As String = dt.Rows(i)("CardCode")
            Dim strPwd As String = dt.Rows(i)("Password")

            sqlCm.Parameters("@CardCode ").Value = cardCode
            sqlCm.Parameters("@Password").Value = oLic.GetEncPwd(strPwd)
            sqlCm.ExecuteNonQuery()

        Next

    End If

	
	sqlCm = New SqlCommand(string.Format("select case when exists(select '' from sysobjects where name = 'DBOLKPostObjectCreation{0}') Then 1 Else 0 End Verfy", dbID), sqlCnCommon)
    
    Dim server As New Server(new ServerConnection(olkip, olklogin, olkpass))
    Dim db As Database = server.Databases("OLKCommon")
    sqlCn.ChangeDatabase("OLKCommon")
    
	If Convert.ToBoolean(sqlCm.ExecuteScalar()) Then
        Dim proc As StoredProcedure = db.StoredProcedures(string.Format("DBOLKPostObjectCreation{0}", dbID))
        
        Dim olkPosStr As String = proc.Script()(2).Replace("CREATE PROCEDURE", "ALTER PROCEDURE")
        
        Dim changed As Boolean = False
        If olkPosStr.IndexOf("@transtype") = -1 Then
        	olkPosStr = olkPosStr.Insert(olkPosStr.ToLower().IndexOf("@object int,"), _
        		String.Format("@transtype char(1),	        -- N = New, U = Update, A = Add{0}@sessiontype char(1),		-- C = Client, A = Agent, P = Pocket{0}", Environment.NewLine))
        	changed = True
        End If
        If olkPosStr.IndexOf("@CurrentSlpCode") = -1 Then
        	Dim indexStr As String = "@lognum int"
        	olkPosStr = olkPosStr.Insert(olkPosStr.ToLower().IndexOf(indexStr)+indexStr.Length, _
        		String.Format(", {0}@CurrentSlpCode int", Environment.NewLine))
        	changed = True
        End If
        If olkPosStr.IndexOf("@Branch") = -1 Then
        	Dim indexStr As String = "@currentslpcode int"
        	olkPosStr = olkPosStr.Insert(olkPosStr.ToLower().IndexOf(indexStr)+indexStr.Length, _
        		String.Format(", {0}@Branch int", Environment.NewLine))
        	changed = True
        End If
        If changed Then
        	sqlCm = New SqlCommand(olkPosStr, sqlCn, sqlTran)
        	sqlCm.ExecuteNonQuery()
        End If
    End If
    
	sqlCm = New SqlCommand(string.Format("select case when exists(select '' from sysobjects where name = 'DBOLKItemInvValCustom{0}') Then 1 Else 0 End Verfy", dbID), sqlCnCommon)

	If Convert.ToBoolean(sqlCm.ExecuteScalar()) Then
        Dim func As UserDefinedFunction = db.UserDefinedFunctions(string.Format("DBOLKItemInvValCustom{0}", dbID))

        Dim olkFunStr As String = func.Script()(2).Replace("CREATE FUNCTION", "ALTER FUNCTION")

        If olkFunStr.IndexOf("@Qty") <> -1 Then
            olkFunStr = Replace(olkFunStr, "@WhsCode nvarchar(8), @Qty numeric(19,6), @cmp nvarchar(254)", "@WhsCode nvarchar(8), @cmp nvarchar(254)")
            sqlCm = New SqlCommand(olkFunStr, sqlCn, sqlTran)
            sqlCm.ExecuteNonQuery()
        End If
    End If

	sqlCn.ChangeDatabase(dbName)
End Sub

Function UpdateDB() As Boolean
	Dim ds as new DataSet()
    Dim sqlDa As New SqlDataAdapter
    Dim sqlStr As string = "select * from olktdb where Deleted = 'N'"
    sqlDa.SelectCommand = New SqlCommand(sqlStr, sqlCnCommon)
    sqlDa.Fill(ds, "olktdb")
 
    Dim LawsSet As String = New SqlCommand("select LawsSet from CINF", sqlCn, sqlTran).ExecuteScalar()
    Select Case LawsSet
    	Case "CL", "CR", "GT", "US", "CA", "BR"
    		LawsSet = "MX"
    	Case "AT", "AU", "BE", "CH", "CZ", "DE", "DK", "ES", "FI", "FR", "CN", "CY", "HU", "IT", "NL", "NO", "PL", "PT", "RU", "SE", "SK", "GB", "ZA"
    		LawsSet = "PA"
	End Select
  
    Dim Language As Integer 
    
    If New SqlCommand("select Left(Version,2) from CINF", sqlCn, sqlTran).ExecuteScalar() = "67" Then
        Language = New SqlCommand("select Case NatLng  " & _
		"	When 'es' Then 2 " & _
		"	When 'fr' Then 8 " & _
		"	When 'pt' Then 6 " & _
		"	When 'he' Then 3 " & _
		"	Else 1 " & _
		"End " & _
		"from OLKCommon ", sqlCn, sqlTran).ExecuteScalar()
    Else
        Language = New SqlCommand("select Case  " & _
		"	When Language = 1 Then 3 " & _
		"	When Language in (23,25) Then 2 " & _
		"	When Language = 19 Then 6 " & _
		"	When Language = 22 Then 8 " & _
		"	Else 1 " & _
		"End Language from CINF", sqlCn, sqlTran).ExecuteScalar()
	End If
    
    sqlStr = "select TableID, PK, Case When Exists( " & _
    "select 'A' from dbo.sysobjects where id = object_id(T0.TableID) and OBJECTPROPERTY(id, N'IsUserTable') = 1) " & _
    "Then 'Y' Else 'N' End As Verfy from olkcommon..olktdb T0 where Type = 'T' and deleted = 'N' and TableID <> 'OLKImgFiles'"
    sqlDa = New SqlDataAdapter()
    sqlDa.SelectCommand = New SqlCommand(sqlStr, sqlCn, sqlTran)
    sqlDa.Fill(ds, "sysobjects")

	Dim sqlCm as sqlCommand
	Dim dr, dr2 as DataRow
    For Each dr In ds.Tables("sysobjects").Rows
        If dr(2) = "Y" Then
            sqlStr = "EXEC sp_rename '" & dr(0) & "', 'TMP" & dr(0) & "'"
            sqlCm = New SqlCommand(sqlStr, sqlCn, sqlTran)
            sqlCm.ExecuteNonQuery()

            If dr(1) = "Y" Then
                sqlStr = "EXEC sp_helpconstraint TMP" & dr(0) & ", nomsg"
                sqlDa.SelectCommand = New SqlCommand(sqlStr, sqlCn, sqlTran)
                sqlDa.Fill(ds, "const")
                If not ds.Tables("const") is nothing Then
	                For Each dr2 In ds.Tables("const").Rows
	                    sqlStr = "alter table TMP" & dr(0) & " drop constraint [" & dr2(1) & "]"
	                    sqlCm = New SqlCommand(sqlStr, sqlCn, sqlTran)
	                    sqlCm.ExecuteNonQuery()
	                Next
	                ds.Tables("const").Clear()
                End If
            End If
            sqlStr = "insert olktempdb (OLDTableID, NEWTableID) Values('" & dr(0) & "', 'TMP" & dr(0) & "')"
            sqlCm = New SqlCommand(sqlStr, sqlCnCommon)
            sqlCm.ExecuteNonQuery()
        End If
    Next

    sqlStr = ""
    Dim dv As DataView = New DataView(ds.Tables("olktdb"))
    Dim drv As DataRowView

    sqlCn.ChangeDatabase("OLKCommon")
    dv.RowFilter = "Type in ('P','F') and TableID not in ('OLKPostObjectCreation', 'OLKSalesItemDetailsCustom', 'OLKItemInvCustom', 'OLKItemInvValCustom', 'OLKItemInvDisValCustom', 'OLKPostAddItemToDoc') "
    For Each drv In dv
        Select Case drv("Type")
            Case "P"
                sqlStr = "if exists (select '' from dbo.sysobjects where id = object_id(N'[dbo].[DB{1}{0}]') and OBJECTPROPERTY(id, N'IsProcedure') = 1) " & _
                "drop procedure [dbo].[DB{1}{0}] "
            Case "F"
                sqlStr = "if exists (select '' from dbo.sysobjects where id = object_id(N'[dbo].[DB{1}{0}]') and xtype in (N'FN', N'IF', N'TF')) " & _
                "drop function [dbo].[DB{1}{0}] "
        End Select
        sqlStr = string.Format(sqlStr, dbID, drv("TableID"))

	    sqlCm = New SqlCommand(sqlStr, sqlCn, sqlTran)
	    sqlCm.ExecuteNonQuery()
    Next
    
    sqlCn.ChangeDatabase(dbName)
    dv.RowFilter = "Deleted = 'N' and Upgrade = 'Y' and Type = 'T' and LawsSet in ('All', '" & LawsSet & "') and NewSAP in ('A','N') and CUFD is null"
    For Each drv In dv
    	lastItem = "Table - " & drv("TableID")
    	sqlStr = drv("Query")
        sqlCm = New SqlCommand(sqlStr, sqlCn, sqlTran)
        sqlCm.ExecuteNonQuery()
    Next
    
    sqlCn.ChangeDatabase("OLKCommon")
    dv.RowFilter = "Deleted = 'N' and Upgrade = 'Y' and Type = 'F' and LawsSet in ('All', '" & LawsSet & "') and NewSAP in ('A','N')"
    For Each drv In dv
    	lastItem = "Function - " & drv("TableID")
        sqlStr = drv("Query")
        sqlCm = New SqlCommand(sqlStr.Replace("{dbID}", dbID).Replace("{dbName}", dbName), sqlCn, sqlTran)
        sqlCm.ExecuteNonQuery()
    Next

    dv.RowFilter = "Deleted = 'N' and Upgrade = 'Y' and Type = 'P' and LawsSet in ('All', '" & LawsSet & "') and NewSAP in ('A','N') and CUFD is null"
    For Each drv In dv
            lastItem = "Procedure - " & drv("TableID")
            sqlStr = drv("Query")
            sqlCm = New SqlCommand(sqlStr.Replace("{dbID}", dbID).Replace("{dbName}", dbName), sqlCn, sqlTran)
            Try
                sqlCm.ExecuteNonQuery()
            Catch ex As Exception

            End Try
    Next
    
    sqlCn.ChangeDatabase(dbName)

    
	Dim varFields, CAdd1, CAdd2, CAdd3 as String
    sqlStr = "select * from olktempdb T0 "
    sqlDa = New SqlDataAdapter
    sqlDa.SelectCommand = New SqlCommand(sqlStr, sqlCnCommon)
    sqlDa.Fill(ds, "olktempdb")
    For Each dr In ds.Tables("olktempdb").Rows
        varFields = ""
        CAdd1 = ""
        CAdd2 = ""
        CAdd3 = ""
        sqlStr = "select top 1 * from " & dr("NEWTableID")
        Dim dsR As New DataSet
        sqlDa = New SqlDataAdapter
        sqlDa.SelectCommand = New SqlCommand(sqlStr, sqlCn, sqlTran)
        sqlDa.Fill(dsR, "oldFields")
        sqlStr = "select name As FieldID from syscolumns where id = object_id('" & dr("NEWTableID") & "') and name not in " & _
                 "(select name from syscolumns where id = object_id('" & dr("OLDTableID") & "'))"
        sqlDa.SelectCommand = New SqlCommand(sqlStr, sqlCn, sqlTran)
        sqlDa.Fill(dsR, "remFields")
        Dim dc as DataColumn
        Dim remField, FoundUserName As Boolean
        For Each dc In dsR.Tables("oldFields").Columns
            remField = False
            If dsR.Tables("remFields").Rows.Count > 0 Then
                For Each dr2 In dsR.Tables("remFields").Rows
                    If dc.ColumnName = dr2("FieldID") Then remField = True
                Next
                If Not remField Then
                    If varFields <> "" Then varFields = varFields & ", "
                    varFields &= String.Format("[{0}]", dc.ColumnName)
                End If
            Else
                If varFields <> "" Then varFields = varFields & ", "
                varFields &= String.Format("[{0}]", dc.ColumnName)
            End If
            If dr("NEWTableID") = "TMPOLKAgentsAccess" And dc.ColumnName = "UserName" Then FoundUserName = True
        Next
        If dr("NEWTableID") = "TMPOLKAgentsAccess" And Not FoundUserName Then
            CAdd1 = "UserName, "
            CAdd2 = "(select SlpName from oslp where slpcode = tmpolkagentsaccess.slpcode), "
        End If
        sqlStr = "insert " & dr("OLDTableID") & "(" & CAdd1 & varFields & ") select " & CAdd2 & varFields & " from " & dr("NEWTableID")
        sqlCm = New SqlCommand(sqlStr, sqlCn, sqlTran)
        sqlCm.CommandTimeout = 1800
        'Try
            sqlCm.ExecuteNonQuery()
        'Catch ex As Exception
        '    ErrMsg &= "Err importing data from " & dr("NEWTableID") & " to " & dr("OLDTableID") & ": " & ex.Message
        '    Err = true
        '    return false
        'End Try

    Next
    
    dv.RowFilter = "Deleted = 'N' and Upgrade = 'Y' and Type = ('Q')"
    For Each drv In dv
    	lastItem = "Query - " & drv("TableID")
        sqlStr = drv("Query")
        sqlCm = New SqlCommand(sqlStr.Replace("{dbID}", dbID).Replace("{dbName}", dbName), sqlCn, sqlTran)
        sqlCm.Parameters.Add("@LanID", SqlDbType.Int).Value = Language
        'Try
            sqlCm.ExecuteNonQuery()
       ' Catch ex As Exception
       '     ErrMsg &= "Err executing query " & drv(1) & ": " & ex.Message
       '     Err = true
       '     return false
       ' End Try

    Next

    ds.Tables("sysobjects").Clear()

    sqlStr = ""
    For Each dr In ds.Tables("olktempdb").Rows
        sqlStr = sqlStr & " drop table " & dr("NEWTableID")
    Next
    sqlCm = New SqlCommand(sqlStr, sqlCn, sqlTran)
    sqlCm.ExecuteNonQuery()
    ds.Tables("olktempdb").Clear()

    sqlStr = "declare @Version nvarchar(100) set @Version = (select Version from OLKCommon..OLKAdminLogin) " & _
    "update olkcommon..olkdba set version = @Version where dbName = '" & Request("dbName") & "' " & _
    "update olkcommon set Version = @Version " & _
    "if not exists(select 'A' from olkdocconf where ObjectCode = 48) " & _
    "insert olkDocConf(ObjectCode, ObjectName, Confirm, Active, Series, T1) " & _
    "values(48, 'Factura Contado', 'N', 'Y', 0, '0') "
    sqlCm = New SqlCommand(sqlStr, sqlCn, sqlTran)
    'Try
        sqlCm.ExecuteNonQuery()
    'Catch ex As Exception
    '    ErrMsg &= "Err finishing " & Request("dbName") & " update: " & ex.Message
    '    Err = true
    '    return false
    'End Try

end function
</script>
<!--#include file="../conn.asp"-->	
<% 
Dim key as RegistryKey 
Try
	key = Registry.LocalMachine.OpenSubKey("Software\Microsoft\MSSQLServer\MSSQLServer\")
	BackupPath = key.GetValue("BackupDirectory")
Catch 
End Try
If BackupPath = "" Then BackupPath = "%temp%"
StartUpdate()
If Not Err Then
	Response.Redirect("../adminSubmit.asp?submitCmd=changeDb&dbName="&dbID)
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
		<%=getLangVal("LttlUpd")%></td>
		<td>
		<p align="center">
		<img src="images/spacer.gif" width="1" height="15" border="0" alt=""></td>
	</tr>
	<tr>
		<td background="images/ventana_r2_c1.gif">
		<p align="center">
		<img name="ventana_r2_c1" src="images/ventana_r2_c1.gif" width="1" height="263" border="0" alt=""></td>
		<td bgcolor="#FFFFFF" background="images/ventana_r2_c2.gif" style="font-size: x-small; font-family: Verdana; color: #3580A8; padding: 10px;"><%=String.Format(getLangVal("LtxtUpdErr"), ErrMsg)%><br/>
		<p align="center"><input type="button" name="btnGoBack" value="<%=getLangVal("LtxtReturn")%>" onclick="javascript:window.location.href='../admin.asp';" style="color: #68A6C0; font-family: Tahoma; border: 1px solid #68A6C0; background-color: #E5F1FF; font-size:10px; height:23; font-weight:bold"></p></td>
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
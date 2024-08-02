<!--#include file="clientInc.asp"-->
<% Select Case userType
Case "C" %><!--#include file="clientTop.asp"-->
<% Case "V" %><!--#include file="agentTop.asp"-->
<% End Select %>
<!--#include file="lang/configErr.asp" -->
<% sql = getConfigErrQry
set rs = conn.execute(sql)
%>
<div align="center">
	<table border="0" id="table1" cellspacing="0" cellpadding="0" width="435">
		<tr>
			<td height="182" background="images/error_olCOnfig.gif" valign="top">
			<table border="0" cellspacing="0" width="100%" id="table2">
				<tr>
					<td height="24">
					<p align="center"><b>
					<font face="Verdana" size="2" color="#0066CC">
					&nbsp;<%=getconfigErrLngStr("LttlConfErr")%></font></b></td>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td style="font-family: Verdana; font-size: 10px; border-left: 1px solid #0066CC; border-right: 1px solid #0066CC; border-bottom: 1px solid #0066CC; padding-right:10px" bgcolor="#F9FCFF">
			<br>
			<ol style="color: #4783C5">
				<% If rs("RestoreUDFErr") = "Y" Then %><li><font color="#4783C5"><b>
				<%=getconfigErrLngStr("LtxtObsRestUDF")%></b></font></li><% End If %>
				<% If rs("OBServerUserErr") = "Y" Then %><li><font color="#4783C5"><b><% If 1 = 2 Then %><%=getconfigErrLngStr("LtxtDBUser")%><% Else %><%=Replace(getconfigErrLngStr("LtxtDBUser"), "{0}", Session("olkdb"))%><% End If %></b></font></li><% End If %>
				<% If rs("OBServerActiveErr") = "Y" Then %><li><font color="#4783C5"><b><% If 1 = 2 Then %><%=getconfigErrLngStr("LtxtDBActive")%><% Else %><%=Replace(getconfigErrLngStr("LtxtDBActive"), "{0}", Session("olkdb"))%><% End If %></b></font></li><% End If %>
				<% If rs("WhsDefErr") = "Y" Then %><li><font color="#4783C5"><b>
				<%=getconfigErrLngStr("LtxtDefWhs")%></b></font></li><% End If %>
				<% If rs("PayAcctErr") = "Y" Then %><li><font color="#4783C5"><b><% If 1 = 2 Then %><%=getconfigErrLngStr("LtxtCondRctAcct")%><% Else %><%=Replace(getconfigErrLngStr("LtxtCondRctAcct"), "{0}", LCase(txtRct))%><% End If %></b></font></li><% End If %>
				<% If rs("OCRDActCurErr") = "Y" Then %><li><font color="#4783C5"><b><% If 1 = 2 Then %><%=getconfigErrLngStr("LtxtAcctCurOCRD")%><% Else %><%=Replace(Replace(Replace(getconfigErrLngStr("LtxtAcctCurOCRD"), "{0}", rs("DebPayAcct")), "{1}", txtClient), "{2}", Session("UserName"))%><% End If %>
				</b></font></li><% End If %>
				<% If rs("CurRateErr") = "Y" Then %><li><font color="#4783C5"><b>
				<%=getconfigErrLngStr("LtxtCurrRate")%></b></font></li><% End If %>
				<% If rs("SeriesErr") = "Y" Then %><li><font color="#4783C5"><b><% If 1 = 2 Then %><%=getconfigErrLngStr("LtxtSeriesConf")%><% Else %><%=Replace(getconfigErrLngStr("LtxtSeriesConf"), "{0}", getConfigErrQryVar("DocDesc"))%><% End If %></b></font></li><% End If %>
				<% If Request("errCmd") = "AsignedSLP" Then
				sql = "select IsNull(T0.SlpName, '') SlpName, IsNull(T1.CardName, '') CardName " & _
				"from OSLP T0 " & _
				"inner join OCRD T1 on T1.SlpCode = T0.SlpCode " & _
				"where T1.CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "'"
				set rs = conn.execute(sql) %><li><font color="#4783C5"><b><% If 1 = 2 Then %><%=getconfigErrLngStr("LtxtErrClientAgent")%><% Else %><%=Replace(Replace(Replace(Replace(getconfigErrLngStr("LtxtErrClientAgent"), "{0}", txtClient), "{1}", rs("CardName")), "{2}", rs("SlpName")), "{3}", txtAgent)%><% End If %></b></font></li><% End If %>
			</ol>
			</td>
		</tr>
	</table>
</div>
<% Function getConfigErrQry
	If Request("errCmd") = "Pay" or Request("errCmd") = "PayDoc" or Request("errCmd") = "Doc" Then
		RetVal = "declare @CardCode nvarchar(15) set @CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "' "
	End If
	Select Case Request("errCmd")
		Case "Doc"
			obj = Request("obj")
			obj2 = Request("obj")
		Case "PayDoc"
			obj = 13
			obj2 = 48
		Case "Pay"
			obj = 24
			obj2 = 24
	End Select
	RetVal = RetVal & "select case when exists( " & _
		"	select 'A' " & _
		"	from OLKCUFD T0 " & _
		"	inner join CUFD T1 on T1.TableID = T0.TableID and T1.FieldID = T0.FieldID " & _
		"	where T0.TableID = '" & getConfigErrQryVar("MainTable") & "' and T0.Active = 'Y' and not exists " & _
		"	(select 'A' from R3_ObsCommon..syscolumns where id =  " & _
		"		(select id from R3_ObsCommon..sysobjects where name = '" & getConfigErrQryVar("OBSTable") & "')  " & _
		"	and name =  " & _
		"	IsNull( " & _
		"		(select SDKID collate database_default from R3_ObsCommon..TCIF where CompanyDB = db_name()),'')  " & _
		"		++ T1.AliasID) " & _
		") "
		
	If Request("errCmd") = "PayDoc" or Request("errCmd") = "Doc" Then
		RetVal = RetVal & " or exists( " & _
				"	select 'A' " & _
				"	from OLKCUFD T0 " & _
				"	inner join CUFD T1 on T1.TableID = T0.TableID and T1.FieldID = T0.FieldID " & _
				"	where T0.TableID = '" & getConfigErrQryVar("MainTable2") & "' and T0.Active = 'Y' and not exists " & _
				"	(select 'A' from R3_ObsCommon..syscolumns where id =  " & _
				"		(select id from R3_ObsCommon..sysobjects where name = '" & getConfigErrQryVar("OBSTable2") & "')  " & _
				"	and name =  " & _
				"	IsNull( " & _
				"		(select SDKID collate database_default from R3_ObsCommon..TCIF where CompanyDB = db_name()),'')  " & _
				"		++ T1.AliasID) " & _
				") "
	ElseIf Request("errCmd") = "Card" Then
		RetVal = RetVal & " or exists( " & _
				"	select 'A' " & _
				"	from OLKCUFD T0 " & _
				"	inner join CUFD T1 on T1.TableID = T0.TableID and T1.FieldID = T0.FieldID " & _
				"	where T0.TableID = 'CRD1' and T0.Active = 'Y' and not exists " & _
				"	(select 'A' from R3_ObsCommon..syscolumns where id =  " & _
				"		(select id from R3_ObsCommon..sysobjects where name = 'CRD1')  " & _
				"	and name =  " & _
				"	IsNull( " & _
				"		(select SDKID collate database_default from R3_ObsCommon..TCIF where CompanyDB = db_name()),'')  " & _
				"		++ T1.AliasID) " & _
				") or exists( " & _
				"	select 'A' " & _
				"	from OLKCUFD T0 " & _
				"	inner join CUFD T1 on T1.TableID = T0.TableID and T1.FieldID = T0.FieldID " & _
				"	where T0.TableID = 'OCPR' and T0.Active = 'Y' and not exists " & _
				"	(select 'A' from R3_ObsCommon..syscolumns where id =  " & _
				"		(select id from R3_ObsCommon..sysobjects where name = 'CRD2')  " & _
				"	and name =  " & _
				"	IsNull( " & _
				"		(select SDKID collate database_default from R3_ObsCommon..TCIF where CompanyDB = db_name()),'')  " & _
				"		++ T1.AliasID) " & _
				") "

	End If
		
		RetVal = RetVal & " Then 'Y' Else 'N' End RestoreUDFErr, "
		
	If Request("errCmd") = "PayDoc" or Request("errCmd") = "Doc" Then
		RetVal = RetVal & "Case When Not exists(select 'A' from owhs where WhsCode = (select WhsCode from olkcommon)) Then 'Y' Else 'N' End WhsDefErr, "
	Else
		RetVal = RetVal & "'N' WhsDefErr, "
	End If
	
	If Request("errCmd") = "PayDoc" or Request("errCmd") = "Pay" or Request("errCmd") = "Doc" Then
	RetVal = RetVal & "Case When Not Exists(select 'A' from nnm1 where ObjectCode = Convert(nvarchar(100)," & obj & ") and Series = " & _
		"(select Series from OLKDocConf where ObjectCode = " & obj2 & ")) Then 'Y' Else 'N' End SeriesErr, "
	

	RetVal = RetVal & "(select Case When T0.Currency <> T1.ActCurr and T1.ActCurr <> '##' Then 'Y' Else 'N' End from OCRD T0 " & _
	"inner join OACT T1 on T1.AcctCode = T0.DebPayAcct " & _
	"where T0.CardCode = N'" &  saveHTMLDecode(Session("UserName"), False) & "') OCRDActCurErr, " & _
	"(select DebPayAcct from OCRD where CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "') DebPayAcct, "
		
	RetVal = RetVal & "(select  " & _
		"case when T0.Currency <> (select top 1 MainCurncy from oadm)  " & _
		"and ( " & _
		"	(T0.Currency <> '##'  " & _
		"	and not exists(select 'A' from ORTT where Currency = T0.Currency and DateDiff(day,getdate(),RateDate) = 0))  " & _
		"	or  " & _
		"	(T0.Currency = '##'  " & _
		"	and exists(select T0.CurrCode from OCRN T0 " & _
		"	left outer join ORTT T1 on T1.Currency = T0.CurrCode and DateDiff(day,getdate(),RateDate) = 0 " & _
		"	where T0.CurrCode <> (select top 1 MainCurncy from oadm) and T1.Currency is null)) " & _
		") Then 'Y' Else 'N' End CurRateErr " & _
		"from ocrd T0 where CardCode = @CardCode) CurRateErr, "
	Else
		RetVal = RetVal & "'N' OCRDActCurErr, 'N' SeriesErr, 'N' CurRateErr, "
	End If
		
	If Request("errCmd") = "PayDoc" Then
	RetVal = RetVal & "Case When Not Exists(select 'A' from nnm1 where ObjectCode = '24' and Series =  " & _
		"(select Series2 from OLKDocConf where ObjectCode = 48)) Then 'Y' Else 'N' End Series2Err, "
	Else
		RetVal = RetVal & "'N' Series2Err, "		
	End If
	
	If Request("errCmd") = "Pay" or Request("errCmd") = "PayDoc" Then
	RetVal = RetVal & "Case When Not Exists(select 'A' from OACT where AcctCode =  " & _
		"(select CashAcct from OLKDocConf where ObjectCode = " & obj2 & ")) or " & _
		"Not Exists(select 'A' from OACT where AcctCode =  " & _
		"(select CheckAcct from OLKDocConf where ObjectCode = " & obj2 & ")) Then 'Y' Else 'N' End PayAcctErr, "
	Else
		RetVal = RetVal & "'N' PayAcctErr, "
	End If
	
	RetVal = RetVal & "Case When not exists(select 'A' from R3_ObsCommon..TCIF where CompanyDB = db_name() and uid is not null) Then 'Y' Else 'N' End OBServerUserErr, " & _
		"Case When (select Active from R3_ObsCommon..TCIF where CompanyDB = db_name()) <> 'Y' Then 'Y' Else 'N' End OBServerActiveErr "

	getConfigErrQry = RetVal
End Function

Function getConfigErrQryVar(Var)
	Select Case Var
		Case "MainTable"
			Select Case Request("errCmd")
				Case "DocLines"
					getConfigErrQryVar = "INV1"
				Case "PayDoc"
					getConfigErrQryVar = "OINV"
				Case "Pay"
					getConfigErrQryVar = "ORCT"
				Case "Item"
					getConfigErrQryVar = "OITM"
				Case "Card"
					getConfigErrQryVar = "OCRD"
				Case "Doc"
					getConfigErrQryVar = "OINV"
			End Select
		Case "OBSTable"
			Select Case Request("errCmd")
				Case "DocLines"
					getConfigErrQryVar = "DOC1"
				Case "PayDoc"
					getConfigErrQryVar = "TDOC"
				Case "Pay"
					getConfigErrQryVar = "TPMT"
				Case "Item"
					getConfigErrQryVar = "TITM"
				Case "Card"
					getConfigErrQryVar = "TCRD"
				Case "Doc"
					getConfigErrQryVar = "TDOC"
			End Select
		Case "MainTable2"
			Select Case Request("errCmd")
				Case "PayDoc"
					getConfigErrQryVar = "ORCT"
				Case "Doc"
					getConfigErrQryVar = "INV1"
			End Select
		Case "OBSTable2"
			Select Case Request("errCmd")
				Case "PayDoc"
					getConfigErrQryVar = "TPMT"
				Case "Doc"
					getConfigErrQryVar = "DOC1"
			End Select
		Case "DocDesc"
			Select Case Request("errCmd")
				Case "PayDoc"
					getConfigErrQryVar = txtInv & "/" & txtRct
				Case "Pay"
					getConfigErrQryVar = txtRct '"Recibo"
				Case "Doc"
					Select Case Request("obj")
						Case 13
							getConfigErrQryVar = txtInv '"Factura"
						Case 17
							getConfigErrQryVar = txtOrdr '"Pedido"
						Case 23
							getConfigErrQryVar = txtQuote '"Cotizaci�n"
					End Select
			End Select
	End Select
End Function
 %>
<% Select Case userType
Case "C" %><!--#include file="clientBottom.asp"-->
<% Case "V" %><!--#include file="agentBottom.asp"-->
<% End Select %>
<!--#include file="lang/carddetails.asp" -->
<%
If Request("cxc") <> "Y" Then 
	sql = 	"select UPPER(CardCode) as CardCode, " & _
			"OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRD', 'CardName', T0.CardCode, T0.CardName) CardName, " & _
			"OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCPR', 'Name', (select CntctCode from OCPR where CardCode = T0.CardCode and Name = T0.CntctPrsn), T0.CntctPrsn) CntctPrsn, " & _
			"OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRG', 'GroupName', T0.GroupCode, T1.GroupName) GroupName, " & _
			"OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRY', 'Name', T0.Country, T2.Name) Name, " & _
			"Phone1, Phone2, Fax, Cellular, E_Mail, Balance, Picture, Notes, Currency, CardType "
			
If not myAut.HasAuthorization(174) Then
	sql = sql & ", Case When T0.SlpCode = " & Session("vendid") & " Then 'Y' Else 'N' End IsBPAssigned "
Else
	sql = sql & ", 'Y' IsBPAssigned "
End If


If Not IsNull(myApp.AgentClientsFilter) Then
	sql = sql & ", Case When T0.CardCode in (" & Replace(Replace(myApp.AgentClientsFilter, "@SlpCode", Session("vendid")), "@Type", 2) & ") Then 'N' Else 'Y' End ShowBalance "
Else
	sql = sql & ", 'Y' ShowBalance "
End If

sql = sql & "from ocrd T0 " & _
	"inner join ocrg T1 on T1.groupcode = T0.groupcode " & _
	"left outer join ocry T2 on T2.code = T0.country " & _
	"where cardcode = N'" & saveHTMLDecode(Request("Card"), False) & "'"
Else
	sql = "select CardType, Balance from OCRD where cardcode = N'" & saveHTMLDecode(Request("Card"), False) & "'"
	set rs = conn.execute(sql)
	
	If rs("CardType") <> "S" Then
		If myApp.SVer >= 8 Then
			colCredit = "BalDueCred"
			colDebit = "BalDueDeb"
		Else
			colCredit = "Credit"
			colDebit = "Debit"
		End If
	Else
		If myApp.SVer >= 8 Then
			colCredit = "BalDueDeb"
			colDebit = "BalDueCred"
		Else
			colCredit = "Debit"
			colDebit = "Credit"
		End If
	End If

	sql = "declare @CardCode nvarchar(15) set @CardCode = N'" & saveHTMLDecode(Request("Card"), False) & "' " & _
	"declare @Credit numeric(19,6) declare @d121 numeric(19,6) declare @d120 numeric(19,6) declare @d90 numeric(19,6) declare @d60 numeric(19,6) declare @d30 numeric(19,6) " & _
	"set @Credit = (select isnull(sum(" & colCredit & "),0) From jdt1 where shortname = @CardCode and refdate <= getdate()) " & _
	"set @d30 = (select isnull(sum(" & colDebit & "),0) From jdt1 where shortname = @CardCode and refdate between DateAdd(day,-30,getdate()) and getdate()) " & _
	"set @d60 = (select isnull(sum(" & colDebit & "),0) From jdt1 where shortname = @CardCode and refdate between DateAdd(day,-60,getdate()) and DateAdd(day,-31,getdate())) " & _
	"set @d90 = (select isnull(sum(" & colDebit & "),0) From jdt1 where shortname = @CardCode and refdate between DateAdd(day,-90,getdate()) and DateAdd(day,-61,getdate())) " & _
	"set @d120 = (select isnull(sum(" & colDebit & "),0) From jdt1 where shortname = @CardCode and refdate between DateAdd(day,-120,getdate()) and DateAdd(day,-91,getdate())) " & _
	"set @d121 = (select isnull(sum(" & colDebit & "),0) From jdt1 where shortname = @CardCode and refdate <= DateAdd(day,-121,getdate())) " & _
	"set @d121 = @d121 - @credit If @d121 < 0 Begin set @credit = @d121 set @d121 = 0 End Else Begin set @Credit = 0 End " & _
	"set @d120 = @d120 + @Credit If @d120 < 0 Begin set @Credit = @d120 set @d120 = 0 End Else Begin set @Credit = 0 End " & _
	"set @d90 = @d90 + @credit If @d90 < 0 Begin set @Credit = @d90 set @d90 = 0 End Else Begin set @Credit = 0 End " & _
	"set @d60 = @d60+ @Credit If @d60 < 0 Begin set @credit = @d60 set @d60 = 0 End Else Begin set @Credit = 0 End " & _
	"set @d30 = @d30 + @credit If @d30 < 0 Begin set @d30 = 0 End " & _
	"select UPPER(CardCode) as CardCode, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OCRD', 'CardName', T0.CardCode, T0.CardName) CardName,  " & _
	"Balance, Picture, Notes, Currency, @d30 'd30', @d60 'd60', @d90 'd90', @d120 'd120', @d121 'd121', " & _
	"CardType "

	If not myAut.HasAuthorization(174) Then
		sql = sql & ", Case When T0.SlpCode = " & Session("vendid") & " Then 'Y' Else 'N' End IsBPAssigned "
	Else
		sql = sql & ", 'Y' IsBPAssigned "
	End If

	If Not IsNull(myApp.AgentClientsFilter) Then
		sql = sql & ", Case When T0.CardCode in (" & Replace(Replace(myApp.AgentClientsFilter, "@SlpCode", Session("vendid")), "@Type", 2) & ") Then 'N' Else 'Y' End ShowBalance "
	Else
		sql = sql & ", 'Y' ShowBalance "
	End If

	
	sql = sql & "from ocrd T0 " & _
	"left outer join ocrg T1 on T1.GroupCode = T0.GroupCode  " & _
	"inner join ocry T2 on T2.Code = T0.Country " & _
	"where Cardcode = @CardCode "
End If
set rs = conn.execute(sql)

If rs("Picture") <> "" Then
	Pic = rs("Picture")
Else
	Pic = "pcard.gif"
End If 
%><div align="center">
   <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
    <tr>
      <td bgcolor="#9BC4FF">
      <table border="0" cellpadding="0" bordercolor="#111111" id="AutoNumber1" width="100%">
        <tr>
          <td>
          <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><% Select Case rs("CardType")
          Case "C" %><%=getcarddetailsLngStr("LtxtClientDetails")%>
          <% Case "L" %><%=getcarddetailsLngStr("LtxtLeadDetails")%>
          <% Case "S" %><%=getcarddetailsLngStr("LtxtProvDetails")%>
          <% End Select %></font></b></td>
        </tr>
        <tr>
          <td valign="top">
          <table border="0" cellpadding="0" cellspacing="1"  bordercolor="#111111" id="AutoNumber2" width="100%">
            <tr>
              <td width="30%" valign="middle">
              <p align="center"><a href="operaciones.asp?cmd=viewImage&amp;FileName=<%=Pic%>">
              <img src="pic.aspx?filename=<%=Pic%>&dbName=<%=Session("olkdb")%>&MaxSize=120" border="0"></a></td>
              <td valign="top" align="center">
              <table border="0" cellpadding="0"  bordercolor="#111111" id="AutoNumber3">
              	<% If rs("CardType") <> "S" and (myApp.EnableORDR or myApp.EnableOQUT) Then %>
                <tr>
                  <td>
                  <p>
                  <a href="operaciones.asp?cmd=docgo&c1=<%=CleanItem(rs("CardCode"))%>">
                  <img border="0" src="images/listapen_icon.gif" align="middle"><font color="#000000" size="1" face="Verdana"><%=getcarddetailsLngStr("DtxtDocs")%></font></a></td>
                </tr>
                <% End If %>
              	<% If myApp.EnableOCLG Then %>
                <tr>
                  <td>
                  <p>
                  <a href="operaciones.asp?cmd=goActivities&CardCode=<%=CleanItem(rs("CardCode"))%>">
                  <img border="0" src="images/listapen_icon.gif" align="middle"><font color="#000000" size="1" face="Verdana"><%=getcarddetailsLngStr("DtxtActivities")%></font></a></td>
                </tr>
                <% End If %>
                <% If myAut.HasAuthorization(66) Then %>
                <tr>
                  <td>
                  <p>
                  <a href="goCrdEdit.asp?CardCode=<%=CleanItem(rs("CardCode"))%>">
                  <img border="0" src="images/modify_icon.gif" align="middle"><font color="#000000" size="1" face="Verdana"><%=getcarddetailsLngStr("LtxtEditData")%></font></a></td>
                </tr>
                <% End If %>
                <tr>
                  <td>
                  <p>
                  <a href="operaciones.asp?cmd=searchclient">
                  <img border="0" src="images/search_icon.gif" align="middle"><font color="#000000" size="1" face="Verdana"><%=getcarddetailsLngStr("LtxtBackToSearch")%></font></a></td>
                </tr>
              </table>
              </td>
            </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td>
          <table border="0" cellpadding="0"  bordercolor="#111111" id="AutoNumber4" width="100%">
            <tr>
              <td bgcolor="#7DB1FF" align="left"><b>
              <font size="1" face="Verdana"><%=getcarddetailsLngStr("DtxtCode")%></font></b></td>
              <td bgcolor="#8CBAFF">
				<table cellpadding="0" cellspacing="2" border="0" width="100%">
					<tr>
		              <td bgcolor="#8CBAFF"><font size="1" face="Verdana"><% = rs("CardCode") %></font></td>
		              <td bgcolor="#8CBAFF"><p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><% If rs("CardType") <> "L" and myAut.HasAuthorization(24) and rs("IsBPAssigned") = "Y" and rs("ShowBalance") = "Y" Then %><a href="operaciones.asp?cmd=datos&card=<%=CleanItem(rs("CardCode"))%>&cxc=Y"><font  size="1" face="Verdana" color="#000000"><nobr><span dir="ltr"><%=myApp.MainCur%>&nbsp;<% = FormatNumber(rs("Balance"),myApp.SumDec) %></span></nobr></font></a><% Else %><font  size="1" face="Verdana" color="#000000">****</font><% End If %></p></td>
					</tr>				
				</table>
              </td>
            </tr>
            <tr>
              <td bgcolor="#7DB1FF" align="left"><b>
              <font size="1" face="Verdana"><%=getcarddetailsLngStr("DtxtName")%></font></b></td>
              <td bgcolor="#8CBAFF"><font size="1" face="Verdana"><% = rs("CardName") %></font></td>
            </tr>
            <% If Request("cxc") <> "Y" Then %>
            <tr>
              <td bgcolor="#7DB1FF" align="left"><b>
              <font size="1" face="Verdana"><%=getcarddetailsLngStr("DtxtPhone")%>&nbsp;1</font></b></td>
              <td bgcolor="#8CBAFF"><font size="1" face="Verdana"><% = rs("Phone1") %></font></td>
            </tr>
            <% if rs("Phone2") <> "" then %>
            <tr>
              <td bgcolor="#7DB1FF" align="left"><b>
              <font size="1" face="Verdana"><%=getcarddetailsLngStr("DtxtPhone")%>&nbsp;2</font></b></td>
              <td bgcolor="#8CBAFF"><font size="1" face="Verdana"><% = rs("Phone2") %></font></td>
            </tr>
            <% end if %>
            <tr>
              <td bgcolor="#7DB1FF" align="left"><b>
              <font size="1" face="Verdana"><%=getcarddetailsLngStr("DtxtFax")%></font></b></td>
              <td bgcolor="#8CBAFF"><font face="Verdana" size="1"><% = rs("Fax") %></font></td>
            </tr>
            <tr>
              <td bgcolor="#7DB1FF" align="left"><b>
              <font size="1" face="Verdana"><%=getcarddetailsLngStr("LtxtMobile")%></font></b></td>
              <td bgcolor="#8CBAFF"><font size="1" face="Verdana"><% = rs("Cellular") %></font></td>
            </tr>
            <tr>
              <td bgcolor="#7DB1FF" align="left"><b>
              <font size="1" face="Verdana"><%=getcarddetailsLngStr("DtxtEMail")%></font></b></td>
              <td bgcolor="#8CBAFF"><font size="1" face="Verdana"><% = rs("E_Mail") %></font></td>
            </tr>
            <tr>
              <td bgcolor="#7DB1FF" align="left"><b>
              <font size="1" face="Verdana"><%=getcarddetailsLngStr("DtxtContact")%></font></b></td>
              <td bgcolor="#8CBAFF"><font size="1" face="Verdana"><% = rs("CntctPrsn") %></font></td>
            </tr>
            <tr>
              <td bgcolor="#7DB1FF" align="left"><b>
              <font size="1" face="Verdana"><%=getcarddetailsLngStr("DtxtGroup")%></font></b></td>
              <td bgcolor="#8CBAFF"><font size="1" face="Verdana"><% = rs("GroupName") %></font></td>
            </tr>
            <tr>
              <td bgcolor="#7DB1FF" align="left"><b>
              <font size="1" face="Verdana"><%=getcarddetailsLngStr("DtxtCountry")%></font></b></td>
              <td bgcolor="#8CBAFF"><font size="1" face="Verdana"><% = rs("Name") %></font></td>
            </tr>
            <% If Not IsNull(rs("Notes")) and rs("Notes") <> "" Then %>
            <tr>
              <td bgcolor="#7DB1FF" align="left" valign="top"><b>
              <font size="1" face="Verdana"><%=getcarddetailsLngStr("DtxtNote")%></font></b></td>
              <td bgcolor="#8CBAFF"><font size="1" face="Verdana"><% = rs("Notes") %></font></td>
            </tr>
            <% End If %>
			<%
			set rx = Server.CreateObject("ADODB.RecordSet")
			set rxVal = Server.CreateObject("ADODB.RecordSet")
			sql = 	"select T0.rowIndex, IsNull(T1.AlterRowName, T0.rowName) rowName, T0.rowField, T0.RowType, T0.RowTypeRnd, T0.RowTypeDec, T0.rowOP " & _
					"from olkcardrep T0 " & _
					"left outer join OLKCardRepAlterNames T1 on T1.rowIndex = T0.rowIndex and T1.LanID = " & Session("LanID") & " " & _
					"where T0.rowAccess in ('T','V') and T0.rowOP in ('T','P') " & _
					"order by T0.rowOrder asc"
			rx.open sql, conn, 3, 1   
			If Rx.RecordCount > 0 Then
				sqlx = 	" declare @SlpCode int set @SlpCode = " & Session("vendid") & _
						" declare @CardCode nvarchar(20) set @CardCode = N'" & saveHTMLDecode(Request("Card"), False) & "'" & _
						" declare @dbName nvarchar(100) set @dbName = N'" & Session("OlkDB") & "' " & _
						" declare @LanID int set @LanID = " & Session("LanID") & " " & _
						" select "
				do  while not rx.eof
					If rx.bookmark > 1 Then sqlx = sqlx & ", "
					If rx("rowTypeRnd") = "Y" Then rowTypeRnd = "Convert(Char(1),Convert(int,(10 * rand())))+ + " Else rowTypeRnd = ""
					If rx("rowType") = "L" or rx("rowType") = "M" or rx("rowType") = "H" Then
						Select Case rx("rowTypeDec")
							Case "S"
								myDec = myApp.SumDec
							Case "P"
								myDec = myApp.PriceDec
							Case "R"
								myDec = myApp.RateDec
							Case "Q"
								myDec = myApp.QtyDec
							Case "%"
								myDec = myApp.PercentDec
							Case "M"
								myDec = myApp.MeasureDec
						End Select
					End If
					If rx("rowType") = "L" Then
						sqlx = sqlx & " OLKCommon.dbo.DBOLKCode" & Session("ID") & "('L'," & rowTypeRnd & "Convert(nvarchar(20),(" & Rx("rowField") & ")), " & myDec & ")" & " As '" & Replace(Rx("rowName"), "'", "''") & "'"
					ElseIf rx("rowType") = "M" Then
						sqlx = sqlx & " OLKCommon.dbo.DBOLKCode" & Session("ID") & "('M'," & rowTypeRnd & "Convert(nvarchar(20),(" & Rx("rowField") & ")), " & myDec & ")" & " As '" & Replace(Rx("rowName"), "'", "''") & "'"
					ElseIf rx("rowType") = "H" Then
						sqlx = sqlx & " OLKCommon.dbo.DBOLKCode" & Session("ID") & "('H'," & rowTypeRnd & "Convert(nvarchar(20),(" & Rx("rowField") & ")), " & myDec & ")" & " As '" & Replace(Rx("rowName"), "'", "''") & "'"
					ElseIf rx("rowType") = "F" Then
						sqlx = sqlx & Rx("rowField") & " As '" & Replace(Rx("rowName"), "'", "''") & "'"
					Else
						sqlx = sqlx & "(" & Rx("rowField") & ") As N'" & Replace(Rx("rowName"), "'", "''") & "'"
					End IF
				rx.movenext
				loop
				sqlx = sqlx & " from OCRD where CardCode = @CardCode"
				sqlx = QueryFunctions(sqlx)
				rxVal.open sqlx, conn, 3, 1
				For each fld in rxVal.Fields
				%>
            <tr>
              <td bgcolor="#7DB1FF" align="left" valign="top"><b>
              <font size="1" face="Verdana"><%=fld.Name%></font></b></td>
              <td bgcolor="#8CBAFF"><font size="1" face="Verdana"><%=fld%></font></td>
            </tr>
		<% 		Next
			End If %>

            <% Else %>
            <tr>
              <td bgcolor="#7DB1FF" align="left"><b>
              <font size="1" face="Verdana">30</font></b></td>
              <td bgcolor="#8CBAFF" align="right" dir="ltr"><font size="1" face="Verdana"><nobr><% = rs("Currency") %>&nbsp;<% = FormatNumber(rs("d30"),myApp.SumDec) %></nobr></font></td>
            </tr>
            <tr>
              <td bgcolor="#7DB1FF" align="left"><b>
              <font size="1" face="Verdana">60</font></b></td>
              <td bgcolor="#8CBAFF" align="right" dir="ltr"><font size="1" face="Verdana"><nobr><% = rs("Currency") %>&nbsp;<% = FormatNumber(rs("d60"),myApp.SumDec) %></nobr></font></td>
            </tr>
            <tr>
              <td bgcolor="#7DB1FF" align="left"><b>
              <font size="1" face="Verdana">90</font></b></td>
              <td bgcolor="#8CBAFF" align="right" dir="ltr"><font size="1" face="Verdana"><nobr><% = rs("Currency") %>&nbsp;<% = FormatNumber(rs("d90"),myApp.SumDec) %></nobr></font></td>
            </tr>
            <tr>
              <td bgcolor="#7DB1FF" align="left"><b>
              <font size="1" face="Verdana">120</font></b></td>
              <td bgcolor="#8CBAFF" align="right" dir="ltr"><font size="1" face="Verdana"><nobr><% = rs("Currency") %>&nbsp;<% = FormatNumber(rs("d120"),myApp.SumDec) %></nobr></font></td>
            </tr>
            <tr>
              <td bgcolor="#7DB1FF" align="left"><b>
              <font size="1" face="Verdana">120+</font></b></td>
              <td bgcolor="#8CBAFF" align="right" dir="ltr"><font size="1" face="Verdana"><nobr><% = rs("Currency") %>&nbsp;<% = FormatNumber(rs("d121"),myApp.SumDec) %></nobr></font></td>
            </tr>
            <% End If %>
          </table>
         </td>
        </tr>
        <tr>
          <td>&nbsp;</td>
        </tr>
      </table>
      </td>
    </tr>
    </table>
  </center>
</div>
<% addLngPathStr = "inv/" %>
<!--#include file="lang/activeitem.asp" -->
<%
sql = 	"select T1.Counted, T1.WasCounted, T0.ItemCode, " & _
		"OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'ItemName', T0.ItemCode, T0.ItemName) ItemName, " & _
		"OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OMRC', 'FirmName', T0.FirmCode, T2.FirmName) Marca, " & _
		"OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITB', 'ItmsGrpNam', T0.ItmsGrpCod, T3.ItmsGrpNam) Grupo, " & _
		"Convert(Decimal(20,2),T1.OnHand) as INV, " & _
		"Convert(Decimal(20,2),(T1.OnHand - T1.IsCommited + T1.onorder)) as Disponible, " & _
		"numinsale, salunitMsr, PicturName " & _
		"from oitm T0 " & _
		"inner join oitw T1 on T1.itemcode = T0.itemcode and T1.WhsCode = '" & Session("Bodega") & "' " & _
		"inner join omrc T2 on T2.firmcode = T0.firmcode " & _
		"inner join oitb T3 on T3.itmsgrpcod = T0.itmsgrpcod " & _
		"where InvntItem = 'Y' and T0.ItemCode = N'" & Request.QueryString("Item") & "' "
set rs = conn.execute(sql)

If rs("PicturName") <> "" Then
	Pic = rs("PicturName")'
Else
	Pic = "n_a.gif"
End If 

%><div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
    <tr>
      <td bgcolor="#9BC4FF">
      <form method="POST" action="inv/activeItemUpdate.asp">
        <table border="0" cellpadding="0" bordercolor="#111111" width="100%" id="AutoNumber1">
        <tr>
          <td width="100%">
          <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><%=getactiveitemLngStr("LtxtItemUpdate")%> 
          </font></b></td>
        </tr>
        <tr>
          <td width="100%">
          <table border="0" cellpadding="0" cellspacing="1"  bordercolor="#111111" width="100%" id="AutoNumber2">
            <tr>
              <td width="33%" valign="middle">
                <p align="center">
                <a href="activeitem.asp?cmd=viewImage&amp;FileName=<%=Pic%>"><img border="0" src="pic.aspx?filename=<%=Pic%>&dbName=<%=Session("olkdb")%>&MaxSize=100"></a></td>
              <td width="46%" valign="top">
              <table border="0" cellpadding="0" cellspacing="2"  bordercolor="#111111" width="100%" id="AutoNumber3">
                <tr>
                  <td width="50%" bgcolor="#75ACFF"><b><font face="Verdana">
      <font size="1"><%=getactiveitemLngStr("LtxtOnHand")%></font><font size="1">:
                  </font></font></b></td>
                  <td width="50%" bgcolor="#75ACFF">
                  <font face="Verdana" size="1"><%=RS("INV")%></font></td>
                </tr>
                <% If rs("WasCounted") = "Y" Then %>
                <tr>
                  <td width="50%" bgcolor="#75ACFF"><b>
                  <font size="1" face="Verdana"><%=getactiveitemLngStr("LtxtLastCounted")%></font></b></td>
                  <td width="50%" bgcolor="#75ACFF">
                  <font face="Verdana" size="1"><%=rs("Counted")%></font></td>
                </tr>
                <% end if %>
                <tr>
                  <td width="39%" bgcolor="#75ACFF"><b>
                  <font size="1" face="Verdana"><%=getactiveitemLngStr("DtxtCounted")%></font></b></td>
                  <td width="61%" bgcolor="#75ACFF">
        <font size="1" face="Verdana">
      <input type="text" name="T1" size="3" value="<%=rs("Counted")%>"></font></td>
                </tr>
                <tr>
                  <td width="100%" colspan="2" bgcolor="#75ACFF"><font size="1" face="Verdana">
                 	<input type="radio" value="N"  <% If rs("WasCounted") = "N" Then Response.write "checked " Else Response.Write "disabled" %> name="R1" id="fp1"><b><label for="fp1"><%=getactiveitemLngStr("LtxtNotCounted")%></label></b></font></td>
                </tr>
                <tr>
                  <td width="100%" colspan="2" bgcolor="#75ACFF"><label for="fp1"><font size="1" face="Verdana">
                  <input type="radio" name="R1" value="Y" <% If rs("WasCounted") = "Y" Then Response.write "checked " %>id="fp2"><label for="fp2"><b><%=getactiveitemLngStr("DtxtCounted")%></b></label></font></td>
                </tr>
                <tr>
                  <td width="100%" colspan="2" bgcolor="#75ACFF">
                  <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
                  <input border="0" src="images/ok_icon.gif" name="I2" type="image"></td>
                </tr>
              </table>
              </td>
            </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td width="100%">
          <table border="0" cellpadding="0" cellspacing="1" bordercolor="#111111" width="100%" id="AutoNumber4">
            <tr>
              <td width="28%" bgcolor="#75ACFF"><b>
              <font size="1" face="Verdana"><%=getactiveitemLngStr("DtxtCode")%></font></b></td>
              <td width="72%"><font size="1" face="Verdana">&nbsp;<%=rs("ItemCode")%></font></td>
            </tr>
            <tr>
              <td width="28%" bgcolor="#75ACFF"><b>
              <font size="1" face="Verdana"><%=getactiveitemLngStr("DtxtDesc")%>&nbsp;1</font></b></td>
              <td width="72%"><font size="1" face="Verdana">&nbsp;<%=rs("ItemName")%></font></td>
            </tr>
			<%
			
			      set rx = Server.CreateObject("ADODB.Recordset")
			      set rxVal = Server.CreateObject("ADODB.Recordset")
			      sql = "select T0.rowIndex, IsNull(T1.alterRowName, T0.rowName) rowName, T0.rowField, T0.rowType, T0.rowTypeRnd, T0.rowTypeDec, T0.HideNull, T0.linkActive, T0.linkObject,  " & _
			      	"T2.rsName, Case When T2.rsIndex is not null Then 'Y' Else 'N' End Verfy " & _
					"from olkitemrep T0 " & _
					"left outer join OLKItemRepAlterNames T1 on T1.rowIndex = T0.rowIndex and T1.LanID = " & Session("LanID") & " " & _
			      	"left outer join OLKRS T2 on T2.rsIndex = T0.linkObject " & _
					"where T0.rowAccess in ('T','V') and T0.rowOP in ('T','P') "
					
			      If Session("username") = "" Then sql = sql & " and Convert(varchar(8000),rowField) not like '%@CardCode%' "
			      
			      If Request("cmd") <> "addcart" Then sql = sql & " and Convert(varchar(8000),rowField) not like '%@Quantity%' and Convert(varchar(8000),rowField) not like '%@Price%' and Convert(varchar(8000),rowField) not like '%@Unit%' "
			      
			      If TreeType <> "S" Then
			      	sql = sql & " and T0.rowIndex <> -1 "
			      End If
			            
			      sql = sql & " order by rowOrder "
			      rx.open sql, conn, 3, 1   
			      If Rx.RecordCount > 0 Then
			      If Session("plist") <> "" Then PriceList = " declare @PriceList int set @PriceList = " & Session("plist")
			      If Session("UserName") <> "" Then CardCode = " declare @CardCode nvarchar(20) set @CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "'"
			      		sqlx = ""
			      		 do while not rx.eof
			      		 	varx = varx + 1
			      		 	rowName = Replace(Rx("rowName"), "'", "''")
			      		 	
			      		 	If PriceList = "" and InStr(rx("rowField"),"Price") <> 0 Then
			      		 	ElseIf CardCode = "" and InStr(rx("rowField"),"CardCode") <> 0 Then
			      		 	Else
			      		 		If sqlx <> "" Then sqlx = sqlx & ", "
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
				      		 	Select Case rx("rowType") 
				      		 		Case "L" 
						      		 	sqlx = sqlx & " OLKCommon.dbo.DBOLKCode" & Session("ID") & "('L'," & rowTypeRnd & "Convert(nvarchar(20),(" & Rx("rowField") & ")), " & myDec & ")" & " As N'" & rowName & "'"
						      		 Case "M" 
						      		 	sqlx = sqlx & " OLKCommon.dbo.DBOLKCode" & Session("ID") & "('M'," & rowTypeRnd & "Convert(nvarchar(20),(" & Rx("rowField") & ")), " & myDec & ")" & " As N'" & rowName & "'"
						      		 Case "H" 
						      		 	sqlx = sqlx & " OLKCommon.dbo.DBOLKCode" & Session("ID") & "('H'," & rowTypeRnd & "Convert(nvarchar(20),(" & Rx("rowField") & ")), " & myDec & ")" & " As N'" & rowName & "'"
						      		 Case "F" 
						      		 	sqlx = sqlx & Rx("rowField") & " As N'" & rowName & "'"
					      		 	 Case Else
						      		 	sqlx = sqlx & "(" & Rx("rowField") & ") As N'" & rowName & "'"
				      		 	End Select
				      		 End If
			      		 rx.movenext
			      		 loop
			      		sql = PriceList & CardCode & _
			      			   " declare @SlpCode int set @SlpCode = " & Session("vendid") & _
			      			   " declare @dbName nvarchar(100) set @dbName = '" & Session("OlkDB") & "'" & _
			      			   " declare @ItemCode nvarchar(20) set @ItemCode = N'" & saveHTMLDecode(Request("item"), False) & "' " & _
			      			   " declare @LanID int set @LanID = " & Session("LanID") & " " & _
							   " declare @WhsCode nvarchar(8) set @WhsCode = OLKCommon.dbo.DBOLKGetWhsCode" & Session("ID") & "(" & Session("branch") & ", " & Session("vendid") & ", @ItemCode) " & _
							   " declare @branchIndex int set @branchIndex = " & Session("branch") & " "
							   
						If Request("cmd") = "addcart" Then
							sql = sql & "declare @Quantity numeric(19,6) set @Quantity = " & getNumeric(Qty) & " " & _
										"declare @Price numeric(19,6) set @Price = " & getNumeric(PriceVal) & " " & _
										"declare @Unit int set @Unit = " & Unit & " "
						End If
							   
			      		sql = sql & " select " & sqlx & " from oitm where itemcode = N'" & saveHTMLDecode(Request("item"), False) & "'"
			      		sql = QueryFunctions(sql)
			      		rxVal.open sql, conn, 3, 1
			      End If
			If rx.recordcount > 0 Then
    		For Each Field in rxVal.Fields
    		rx.Filter = "rowName = '" & Field.Name & "'"
    		customVal = rxVal(Field.name)
    		If Not IsNull(customVal) or IsNull(customVal) and rx("HideNull") = "N" Then %>
            <tr>
              <td bgcolor="#75ACFF" align="left" valign="top">
		        <table cellpadding="0" cellspacing="0" border="0" width="100%">
		        	<tr>
		        		<td><font size="1" face="Verdana"><b><%=Field.name%></b></font></td>
		        		<% If rx("linkActive") = "Y" Then %><td width="15">
		        		<a href="javascript:<% If rx("Verfy") = "Y" Then %>doLink(<%=rx("rowIndex")%>)<% Else %>doErrRep()<% End If %>;">
						<img alt="<%=myHTMLEncode(rx("rsName"))%>" border="0" src="images/<%=Session("rtl")%>flechaselec.gif" width="15" height="13" style="cursor: hand"></a></td><% End If %>
		        	</tr>
		        </table>
              </td>
              <td><font size="1" face="Verdana"><%=customVal%></font></td>
            </tr>
  			<% 
  			End If
  			Next
  	 		end if %>

            </table>
           </td>
        </tr>
        </table>
        <input type="hidden" name="item" value="<%=Request("Item")%>">
      </form>
      </td>
    </tr>
    </table>
  </center>
</div>

<%
set rx = Server.CreateObject("ADODB.RecordSet")
set rxVal = Server.CreateObject("ADODB.RecordSet")

sql = "select T0.rowIndex, T1.rsIndex " & _  
"from OLKItemRep T0 " & _  
"inner join OLKRS T1 on T1.rsIndex = T0.linkObject " & _  
"where T0.rowAccess in ('T','V') and T0.rowOP in ('T','P') and linkActive = 'Y' " 

rx.open sql, conn, 3, 1

sql = "select T0.rowIndex, T1.varIndex, T1.varVar, T1.varDataType, T2.valBy, T2.valValue, T2.valDate " & _  
"from OLKItemRep T0 " & _  
"inner join OLKRSVars T1 on T1.rsIndex = T0.linkObject " & _  
"left outer join OLKItemRepLinksVars T2 on T2.rowIndex = T0.rowIndex and T2.varId = T1.varVar " & _  
"where T0.rowAccess in ('T','V') and T0.rowOP in ('T','P') and linkActive = 'Y' " 
rxVal.open sql, conn, 3, 1

%>
<script language="javascript">
<% If Not rx.eof Then %>
function doLink(rowIndex)
{
	switch (rowIndex)
	{
		<% do while not rx.eof %>
		case <%=rx("rowIndex")%>:
			document.frmLink<%=Replace(rx("rowIndex"), "-", "_")%>.submit();
			break;
		<% rx.movenext
		loop %>
	}
}
function doErrRep()
{
	alert('|L:txtErrRep|');
}
<% rx.movefirst
End If %>
function goCartFilter()
{
	document.frmGoCartFilter.submit();
}
<% If Request("cmd") = "addcart" Then %>
setLineTotal()
<% End If %>
</script>
<% do while not rx.eof %>
<form name="frmLink<%=Replace(rx("rowIndex"), "-", "_")%>" id="frmLink<%=Replace(rx("rowIndex"), "-", "_")%>" action="operaciones.asp" method="post">
<input type="hidden" name="rsIndex" value="<%=rx("rsIndex")%>">
<input type="hidden" name="itemCmd" value="A">
<input type="hidden" name="cmd" value="viewRep">
<%
rxVal.Filter = "rowIndex = " & rx("rowIndex")
do while not rxVal.eof
Select Case rxVal("valBy") 
	Case "V"
		If rxVal("varDataType") <> "datetime" Then
			strVal = rxVal("valValue")
		Else
			strVal = rxVal("valDate")
		End If
	Case "F"
		Select Case rxVal("valValue")
			Case "@PriceList"
				strVal = Session("plist")
			Case "@SlpCode"
				strVal = Session("vendid")
			Case "@CardCode"
				strVal = saveHTMLDecode(Session("UserName"), False)
			Case "@WhsCode"
				strVal = saveHTMLDecode(WhsCode, False)
			Case "@dbName"
				strVal = Session("olkdb")
			Case "@ItemCode"
				strVal = saveHTMLDecode(Request("item"), False)
			Case "@Quantity"
				strVal = getNumeric(Qty)
			Case "@Unit"
				strVal = Unit
			Case "@Price"
				strVal = getNumeric(PriceVal)
		End Select
	Case "Q"
		If Session("plist") <> "" Then PriceList = " declare @PriceList int set @PriceList = " & Session("plist")
      	sql = PriceList & _
      			   " declare @SlpCode int set @SlpCode = " & Session("vendid") & _
      			   " declare @CardCode nvarchar(20) set @CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "'" & _
      			   " declare @WhsCode nvarchar(8) set @WhsCode = N'" & saveHTMLDecode(WhsCode, False) & "'" & _
      			   " declare @dbName nvarchar(100) set @dbName = N'" & Session("OlkDB") & "' " & _
      			   " declare @ItemCode nvarchar(20) set @ItemCode = '" & saveHTMLDecode(Request("item"), False) & "' " & _
      			   " select (" & rxVal("valValue") & ") "
		set rv = conn.execute(sql)
		If Not rv.Eof Then strVal = rv(0) Else strVal = ""
End Select
 %>
<input type="hidden" name="var<%=rxVal("varIndex")%>" value="<%=myHTMLEncode(strVal)%>">
<% rxVal.movenext
loop %>
</form>
<% rx.movenext
loop
set rxVal = nothing
set rx = nothing %>
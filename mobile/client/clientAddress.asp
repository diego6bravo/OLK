
<head>
<style type="text/css">
.style1 {
				text-align: center;
}
</style>
</head>

<% addLngPathStr = "client/" %><!--#include file="lang/clientAddress.asp" -->
<%
sql = "select CardCode, IsNull(CardName, '') CardName, " & _
"(select Command from R3_ObsCommon..TLOG where LogNum = T0.LogNum) Command, null nullFldVal " & _
"from R3_ObsCommon..TCRD T0 " & _
"where T0.LogNum = " & Session("CrdRetVal")
set rs = conn.execute(sql)
isUpdate = rs("Command") = "U"
CardCode = rs("CardCode")
CardName = rs("CardName")
fldVal = rs("nullFldVal")

set rg = Server.CreateObject("ADODB.RecordSet")
sql = "select T0.GroupID, IsNull(T1.AlterGroupName, T0.GroupName) GroupName " & _
	"from OLKCUFDGroups T0 " & _
	"left outer join OLKCUFDGroupsAlterNames T1 on T1.TableID = T0.TableID and T1.GroupID = T0.GroupID and T1.LanID = " & Session("LanID") & " " & _
	"where T0.TableID = 'CRD1' and exists(select '' from CUFD X0 left outer join OLKCUFD X1 on X1.TableID = X0.TableID and X1.FieldID = X0.FieldID where X0.TableID = T0.TableID and IsNull(X1.GroupID, -1) = T0.GroupID and AType in ('" & userType & "', 'T') and OP in ('T','P') and Active = 'Y') " & _
	"order by T0.[Order] "
set rg = conn.execute(sql)

set rcOpt = Server.CreateObject("ADODB.RecordSet")
set rcOptVal = Server.CreateObject("ADODB.RecordSet")
sql = "select IsNull(T1.GroupID, -1) GroupID, T0.FieldID, AliasID, IsNull(alterDescr, Descr) Descr, TypeID, SizeID, Dflt, NotNull, IsNull(T1.Pos, 'D') Pos, RTable, " & _
	"Case When Exists(select 'A' from UFD1 where TableId = T0.TableId and FieldId = T0.FieldId) " & _
	"Then 'Y' Else 'N' End As DropDown, NullField, Query, " & _
	"(select SDKID collate database_default from r3_obscommon..tcif where companydb = '" & Session("OlkDB") & "')++AliasID As InsertID " & _
	"from cufd T0 " & _
	"left outer join OLKCUFD T1 on T1.TableID = T0.TableID and T1.FieldID = T0.FieldID " & _
	"left outer join OLKCUFDAlterNames T2 on T2.TableID = T0.TableID and T2.FieldID = T0.FieldID and T2.LanID = " & Session("LanID") & " " & _
	"where T0.TableId = 'CRD1' and AType in ('" & userType & "', 'T') and OP in ('T','P') and Active = 'Y' " & _
	"order by IsNull(T1.GroupID, -1), IsNull(T1.Pos, 'D'), IsNull(T1.[Order], 32727) "
rcOpt.open sql, conn, 3, 1

If Request.Form.Count = 0 Then
	If Request("EditID") <> "" Then
		sqlAddStr = ""
		do while not rcOpt.eof
			sqlAddStr = sqlAddStr & ", "
			sqlAddStr = sqlAddStr & "T0." & rcOpt("InsertID")
		rcOpt.movenext
		loop
		If rcOpt.recordcount > 1 Then rcOpt.movefirst
		
		sql = 	"select T0.NewAddress, T0.Street, T0.Block, T0.City, T0.ZipCode, T0.County, T0.Country, T0.State, T0.TaxCode, Case T0.AdresType " & _
				"	When 'S' Then " & _
				"		Case When T0.NewAddress = T1.ShipToDef Then 'Y' Else 'N' End " & _
				"	When 'B' Then " & _
				"		Case When T0.NewAddress = T1.BillToDef Then 'Y' Else 'N' End " & _
				"	End IsDefault, T0.AdresType" & sqlAddStr & " " & _
				"from R3_ObsCommon..CRD1 T0 " & _
				"inner join R3_ObsCommon..TCRD T1 on T1.LogNum = T0.LogNum " & _
				"where T0.LogNum = " & Session("CrdRetVal") & " and T0.LineNum = " & Request("EditID")
		set rs = conn.execute(sql)
		NewAddress = rs("NewAddress")
		Street = rs("Street")
		Block = rs("Block")
		City = rs("City")
		ZipCode = rs("ZipCode")
		County = rs("County")
		Country = rs("Country")
		State = rs("State")
		TaxCode = rs("TaxCode")
		AdresType = rs("AdresType")
		IsDefault = rs("IsDefault") = "Y"
	Else
		AdresType = Request("AdresType")
	End If
Else
	NewAddress = Request("NewAddress")
	Street = Request("Street")
	Block = Request("Block")
	City = Request("City")
	ZipCode = Request("ZipCode")
	County = Request("County")
	Country = Request("Country")
	State = Request("State")
	TaxCode = Request("TaxCode")
	AdresType = Request("AdresType")
	IsDefault = Request("SetDef") = "Y"
End If
 %>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
  <form method="post" action="client/submitClient.asp" name="frmAdd">
    <tr>
      <td>
      <img src="images/spacer.gif" width="100%" height="1" border="0" alt></td>
    </tr>
    <tr>
      <td bgcolor="#9BC4FF">
      <table border="0" cellpadding="0" bordercolor="#111111" width="100%">
        <tr>
          <td width="100%">
          <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><% If Not isUpdate Then %><%=getclientAddressLngStr("LttlNewClient")%><% Else %><%=getclientAddressLngStr("LttlEditClient")%><% End If %> - <% If Request("EditID") = "" Then %><%=getclientAddressLngStr("LtxtNewAddress")%><% Else %><%=getclientAddressLngStr("LtxtEditAddress")%><% End If %>
          </font></b></td>
        </tr>
        <tr>
          <td width="100%">
         <!--#include file="clientMenu.asp"--></td>
        </tr>
        <tr>
          <td width="100%">
            <table border="0" cellpadding="0"  bordercolor="#111111" width="100%" id="AutoNumber3">
            <tr>
              <td width="30%" bgcolor="#7DB1FF"><b>
              <font size="1" face="Verdana"><%=getclientAddressLngStr("DtxtCode")%></font></b></td>
              <td width="70%" bgcolor="#8CBAFF"><p>
              <font size="1" face="Verdana"><%=myHTMLEncode(CardCode)%></font></td>
            </tr>
			<tr>
              <td width="30%" bgcolor="#7DB1FF"><b>
              <font size="1" face="Verdana"><%=getclientAddressLngStr("DtxtName")%></font></b></td>
              <td width="70%" bgcolor="#8CBAFF"><p>
              <font size="1" face="Verdana"><%=myHTMLEncode(CardName)%></font></td>
            </tr>
            </table>
           </td>
        </tr>
        <tr>
          <td width="100%">
            <table border="0" cellpadding="0"  bordercolor="#111111" width="100%">

					<tr>
						<td width="30%" bgcolor="#7DB1FF"><b>
						<font size="1" face="Verdana"><%=getclientAddressLngStr("DtxtName")%></font></b></td>
						<td>
						<input type="text" name="NewAddress" size="40" maxlength="50" value="<%=myHTMLEncode(NewAddress)%>"></td>
					</tr>
					<tr>
						<td width="30%" bgcolor="#7DB1FF"><b>
						<font size="1" face="Verdana"><%=getclientAddressLngStr("LtxtStreet")%></font></b></td>
						<td>
						<input type="text" name="Street" maxlength="100" style="width: 100%" value="<%=myHTMLEncode(Street)%>"></td>
					</tr>
					<tr>
						<td width="30%" bgcolor="#7DB1FF"><b>
						<font size="1" face="Verdana"><%=getclientAddressLngStr("LtxtBlock")%></font></b></td>
						<td>
						<input type="text" name="Block" maxlength="100" style="width: 100%" value="<%=myHTMLEncode(Block)%>"></td>
					</tr>
					<tr>
						<td width="30%" bgcolor="#7DB1FF"><b>
						<font size="1" face="Verdana"><%=getclientAddressLngStr("DtxtCity")%></font></b></td>
						<td>
						<input type="text" name="City" maxlength="100" style="width: 100%" value="<%=myHTMLEncode(City)%>"></td>
					</tr>
					<tr>
						<td width="30%" bgcolor="#7DB1FF"><b>
						<font size="1" face="Verdana"><%=getclientAddressLngStr("LtxtPostalCode")%></font></b></td>
						<td>
						<input type="text" name="ZipCode" size="20" maxlength="20" value="<%=myHTMLEncode(ZipCode)%>"></td>
					</tr>
					<tr>
						<td width="30%" bgcolor="#7DB1FF"><b>
						<font size="1" face="Verdana"><%=getclientAddressLngStr("LtxtCounty")%></font></b></td>
						<td>
						<p align="center">
						<input type="text" name="County" maxlength="100" style="width: 100%" value="<%=myHTMLEncode(County)%>"></p>
						</td>
					</tr>
					<tr>
						<td width="30%" bgcolor="#7DB1FF"><b>
						<font size="1" face="Verdana"><%=getclientAddressLngStr("DtxtCountry")%></font></b></td>
						<td>
						<select size="1" name="Country" onchange="changeCountry()" style="width: 50%; ">
						<option value=""></option>
						<%
						set cmd = Server.CreateObject("ADODB.Command")
						cmd.ActiveConnection = connCommon
						cmd.CommandType = &H0004
						cmd.CommandText = "DBOLKGetCountries" & Session("ID")
						cmd.Parameters.Refresh()
						cmd("@LanID") = Session("LanID")
						set rd = cmd.execute()
						do while not rd.eof %>
						<option <% If Country = rd("Code") Then %>selected<% End If %> value="<%=rd("Code")%>"><%=rd("Name")%></option>
						<% rd.movenext
						loop %>
						</select></td>
					</tr>
					<tr>
						<td width="30%" bgcolor="#7DB1FF"><b>
						<font size="1" face="Verdana"><%=getclientAddressLngStr("DtxtState")%></font></b></td>
						<td>
						<select size="1" name="State" style="height: 16px; width: 50%; ">
						<option value=""></option>
						<% If Country <> "" Then
							set cmd = Server.CreateObject("ADODB.Command")
							cmd.ActiveConnection = connCommon
							cmd.CommandType = &H0004
							set rd = Server.CreateObject("ADODB.RecordSet")
							cmd.CommandText = "DBOLKGetCountryStates" & Session("ID")
							cmd.Parameters.Refresh()
							cmd("@LanID") = Session("LanID")
							cmd("@Code") = Country
							rd.open cmd, , 3, 1
							do while not rd.eof %>
							<option <% If State = rd("Code") Then %>selected<% End If %> value="<%=rd("Code")%>"><%=rd("Name")%></option>
							<% rd.movenext
							loop
							End If %>
						</select></td>
					</tr>
					<% 
					
					Select Case myApp.LawsSet
						Case "MX", "CL", "CR", "GT", "US", "CA", "BR"
							If Request("AdresType") = "S" Then %>
					<tr>
						<td width="30%" bgcolor="#7DB1FF"><b>
						<font size="1" face="Verdana"><%=getclientAddressLngStr("LtxtTaxCode")%></font></b></td>
						<td>
						<select size="1" name="TaxCode" style="height: 16px; width: 50%; ">
						<option value=""></option>
						<%  sql = "select Code, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OSTC', 'Name', Code, Name) Name from OSTC where ValidForAR = 'Y'"
							set rd = conn.execute(sql)
							do while not rd.eof %>
							<option <% If TaxCode = rd("Code") Then %>selected<% End If %> value="<%=rd("Code")%>"><%=rd("Name")%></option>
							<% rd.movenext
							loop %>
						</select></td>
					</tr>
					<% 	End If
					End Select %>
					<tr>
						<td>&nbsp;</td>
						<td>
						<input type="checkbox" name="SetDef" id="SetDef" <% If isDefault Then %>checked disabled<% End If %> value="Y" style="background: background-image; border: 0px solid"><label for="SetDef"><%=getclientAddressLngStr("LtxtSetAsDef")%></label></td>
					</tr>
					<% do while not rg.eof %>
					<tr>
						<td bgcolor="#7DB1FF" colspan="2"><b>
						<font size="1" face="Verdana"><% Select Case CInt(rg("GroupID"))
						Case -1 %><%=getclientAddressLngStr("DtxtUDF")%><%
						Case Else
							Response.Write rg("GroupName")
						End Select %></font></b></td>
					</tr>
					<% rcOpt.Filter = "GroupID = " & rg("GroupID")
					do while not rcOpt.eof 
                    AliasID = rcOpt("InsertID")
                    If Request.Form.Count = 0 Then
	                    If Request("EditID") <> "" Then fldVal = rs(AliasID)
	                    If rcOpt("TypeID") = "D" Then fldVal = FormatDate(fldVal, False)
	                Else
	                	fldVal = Request("U_" & rcOpt("AliasID"))
	                End If %>
					<tr>
					  <td width="30%" bgcolor="#7DB1FF"><b><font size="1" face="Verdana">&nbsp;<%=rcOpt("Descr")%><% If rcOpt("NullField") = "Y" Then %><font color="red">*</font><% End If %></font></b></td>
					  <td>
						<% If rcOpt("DropDown") = "Y" or Not IsNull(rcOpt("RTable")) then 
						If rcOpt("DropDown") = "Y" Then
							sql = "select FldValue, IsNull(AlterDescr, Descr) Descr " & _
											"from UFD1 T0 " & _
											"left outer join OLKUFD1AlterNames T1 on T1.TableID = T0.TableID and T1.FieldID = T0.FieldID and T1.IndexID = T0.IndexID and T1.LanID = " & Session("LanID") & " " & _
											"where T0.tableid = 'OCLG' and T0.FieldId = " & rcOpt("FieldId")
						Else
							sql = "select Code FldValue, Name Descr from [@" & rcOpt("RTable") & "] order by 2"
						End If
						set rcOptVal = conn.execute(sql) %>
						<font color="#4783C5">
						<select size="1" name="U_<%=rcOpt("AliasID")%>" class="input" style="font-size:10px; width:100%; font-family:Verdana">
						<option></option>
						<% do while not rcOptVal.eof %>
						<option value="<%=rcOptVal("FldValue")%>" <% If fldVal = rcOptVal("FldValue") Then %>selected<% ElseIf rcOptVal("FldValue") = rcOpt("Dflt") and IsNull(fldVal) Then %>selected<% End If %>><%=rcOptVal("Descr")%></option>
						<% rcOptVal.movenext
						loop %></select></font>
						<% Else %>
						<% If rcOpt("Query") = "Y" or rcOpt("TypeID") = "D" Then %><table width="100%" cellspacing="0" cellpadding="0"><tr><td width="16"><a href="#" <% If rcOpt("TypeID") = "D" Then %>onclick="javascript:getCal('<%=rcOpt("AliasID")%>')"<% End If %> <% If rcOpt("Query") = "Y" Then %>onclick="javascript:getVal('<%=rcOpt("AliasID")%>')"<% End If %>><img border="0" src="<% If rcOpt("Query") = "Y" Then %>../images/flechaselec2.gif<% Else %>images/cal.gif<% End If %>"></a></td><td><% End If %>
						<input <% If rcOpt("Query") = "Y" or rcOpt("TypeID") = "D" Then %>readonly<% End If %> type="text" name="U_<%=rcOpt("AliasID")%>" size="<% If rcOpt("TypeID") = "A" Then %>43<% Else %>12<% End If %>" class="input" value="<% If fldVal <> "" Then %><%=fldVal%><% Else %><%=rcOpt("Dflt")%><% End If %>" onclick="this.selectionStart=0;this.selectionEnd=this.value.length;" style="width: 100%; font-family: Verdana; font-size: 10px">
						<% If rcOpt("Query") = "Y" or rcOpt("TypeID") = "D" Then %></td><td width="16"><a href="#" onclick="javascript:document.frmAdd.U_<%=rcOpt("AliasID")%>.value = ''"><img border="0" src="../images/remove.gif" width="16" height="16"></a></td></tr></table><% End If %><% End If %></td>
					</tr>
                    <% 
                    rcOpt.movenext
                    loop 
                    rg.movenext
                    loop
                    rcOpt.Filter = "" %>
					<tr>
						<td colspan="2" class="style1">
						<input type="submit" value="<%=getclientAddressLngStr("DtxtSave")%>" name="btnSave" onclick="return valFrm();"></td>
					</tr>
          </table>
           </td>
        </tr>
        <tr>
          <td width="100%">&nbsp;</td>
        </tr>
      </table>
      </td>
    </tr>
	<input type="hidden" name="cmd" value="address">
	<input type="hidden" name="EditID" value='<%=Request("EditID")%>'>
	<input type="hidden" name="AdresType" value='<%=AdresType%>'>
	<input type="hidden" name="editVar" value="">
	<input type="hidden" name="returnCmd" value="newClientAddress">
	</form>
    </table>
  </center>
</div>
<script type="text/javascript">
function getCal(AliasID)
{
	document.frmAdd.action = 'operaciones.asp';
	document.frmAdd.editVar.value = AliasID;
	document.frmAdd.cmd.value = 'UDFCal';
	document.frmAdd.submit();
}
function getVal(AliasID)
{
	document.frmAdd.action = 'operaciones.asp';
	document.frmAdd.editVar.value = AliasID;
	document.frmAdd.cmd.value = 'UDFQry';
	document.frmAdd.submit();
}
function changeCountry()
{
	document.frmAdd.action = 'operaciones.asp';
	document.frmAdd.cmd.value = 'newClientAddress';
	document.frmAdd.submit();
}
function valFrm()
{
	if (document.frmAdd.NewAddress.value == '')
	{
		alert('<%=getclientAddressLngStr("LtxtValNam")%>');
		document.frmAdd.NewAddress.focus();
		return false;
	}
	<%
	If rcOpt.recordcount > 0 Then rcOpt.movefirst
	do while not rcOpt.eof
	If rcOpt("NullField") = "Y" or rcOpt("TypeID") = "N" or rcOpt("TypeID") = "B" Then
		If rcOpt("NullField") = "Y" Then %>
		if (document.frmAdd.U_<%=rcOpt("AliasID")%>.value == '') 
		{
			alert('<%=getclientAddressLngStr("LtxtValFld")%>'.replace('{0}', '<%=Replace(rcOpt("Descr"), "'", "\'")%>'));
			document.frmAdd.U_<%=rcOpt("AliasID")%>.focus
			return false; 
		}
		<% End If
		If rcOpt("TypeID") = "B" or rcOpt("TypeID") = "N" Then %>
		if (document.frmAdd.U_<%=rcOpt("AliasID")%>.value != '') 
		{
			if (!MyIsNumeric(document.frmAdd.U_<%=rcOpt("AliasID")%>.value)) 
			{
				alert('<%=getclientAddressLngStr("DtxtValNumVal")%>');
				document.frmAdd.U_<%=rcOpt("AliasID")%>.focus
				return false; 
			}
		}
		<% End If
	End If
	rcOpt.movenext
	loop
	rcOpt.close  %>
	document.frmAdd.action = 'client/submitClient.asp';
	document.frmAdd.cmd.value = 'address';
	return true;
}
</script>
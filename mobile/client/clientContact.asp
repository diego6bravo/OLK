
<head>
<style type="text/css">
.style1 {
				text-align: center;
}
</style>
</head>

<% addLngPathStr = "client/" %><!--#include file="lang/clientContact.asp" -->
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
	"where T0.TableID = 'OCPR' and exists(select '' from CUFD X0 left outer join OLKCUFD X1 on X1.TableID = X0.TableID and X1.FieldID = X0.FieldID where X0.TableID = T0.TableID and IsNull(X1.GroupID, -1) = T0.GroupID and AType in ('" & userType & "', 'T') and OP in ('T','P') and Active = 'Y') " & _
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
	"where T0.TableId = 'OCPR' and AType in ('" & userType & "', 'T') and OP in ('T','P') and Active = 'Y' " & _
	"order by IsNull(T1.GroupID, -1), IsNull(T1.Pos, 'D'), IsNull(T1.[Order], 32727) "
rcOpt.open sql, conn, 3, 1

If Request.Form.Count = 0 Then
	If Request("EditID") <> "" Then
		sqlAddStr = ""
		do while not rcOpt.eof
			sqlAddStr = sqlAddStr & ", "
			sqlAddStr = sqlAddStr & rcOpt("InsertID")
		rcOpt.movenext
		loop
		If rcOpt.recordcount > 1 Then rcOpt.movefirst
		
		sql =	"select T0.NewName, T0.Title, T0.Position, T0.Address, T0.tel1, T0.tel2, T0.Cellolar, T0.fax, T0.E_MailL, T0.Pager, T0.Notes1, T0.Notes2, T0.Password, " & _
		"T0.BirthPlace, T0.BirthDate, T0.Gender, T0.Profession, " & _
		"Case When T0.NewName = (select CntctPrsn from R3_ObsCommon..TCRD where LogNum = T0.LogNum) Then 'Y' Else 'N' End IsDefault" & sqlAddStr & " " & _
		"from R3_ObsCommon..CRD2 T0 " & _
		"where T0.LogNum = " & Session("CrdRetVal") & " and T0.LineNum = " & Request("EditID")
		set rs = conn.execute(sql)
		NewName = rs("NewName")
		Title = rs("Title")
		Position = rs("Position")
		Address = rs("Address")
		Tel1 = rs("tel1")
		Tel2 = rs("tel2")
		Cellolar = rs("Cellolar")
		Fax = rs("Fax")
		EMail = rs("E_MailL")
		Pager = rs("Pager")
		Notes1 = rs("Notes1")
		Notes2 = rs("Notes2")
		Password = rs("Password")
		BirthPlace = rs("BirthPlace")
		BirthDate = FormatDate(rs("BirthDate"), False)
		Gender = rs("Gender")
		Profession = rs("Profession")
		isDefault = rs("IsDefault") = "Y"
	End If
Else
	NewName = Request("NewName")
	Title = Request("Title")
	Position = Request("Position")
	Address = Request("Address")
	Tel1 = Request("tel1")
	Tel2 = Request("tel2")
	Cellolar = Request("Cellolar")
	Fax = Request("Fax")
	EMail = Request("E_MailL")
	Pager = Request("Pager")
	Notes1 = Request("Notes1")
	Notes2 = Request("Notes2")
	Password = Request("Password")
	BirthPlace = Request("BirthPlace")
	BirthDate = Request("BirthDate")
	Gender = Request("Gender")
	Profession = Request("Profession")	
	isDefault = Request("SetDef") = "Y"
End If
 %>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
  <form method="post" action="client/submitClient.asp" name="frmCrd">
    <tr>
      <td>
      <img src="images/spacer.gif" width="100%" height="1" border="0" alt></td>
    </tr>
    <tr>
      <td bgcolor="#9BC4FF">
      <table border="0" cellpadding="0" bordercolor="#111111" width="100%" id="AutoNumber1">
        <tr>
          <td width="100%">
          <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><% If Not isUpdate Then %><%=getclientContactLngStr("LttlNewClient")%><% Else %><%=getclientContactLngStr("LttlEditClient")%><% End If %> - <% If Request("EditID") = "" Then %><%=getclientContactLngStr("LtxtNewContact")%><% Else %><%=getclientContactLngStr("LtxtEditContact")%><% End If %>
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
              <font size="1" face="Verdana"><%=getclientContactLngStr("DtxtCode")%></font></b></td>
              <td width="70%" bgcolor="#8CBAFF"><p>
              <font size="1" face="Verdana"><%=myHTMLEncode(CardCode)%></font></td>
            </tr>
			<tr>
              <td width="30%" bgcolor="#7DB1FF"><b>
              <font size="1" face="Verdana"><%=getclientContactLngStr("DtxtName")%></font></b></td>
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
              			<font size="1" face="Verdana"><%=getclientContactLngStr("DtxtName")%></font></b></td>
						<td height="18" class="GeneralTbl">
						<input type="text" name="NewName" size="40" maxlength="50" value="<%=myHTMLEncode(NewName)%>"></td>
					</tr>
					<tr>
						<td width="30%" bgcolor="#7DB1FF"><b>
              			<font size="1" face="Verdana"><%=getclientContactLngStr("LtxtTitle")%></font></b></td>
						<td height="18" class="GeneralTbl">
						<input type="text" name="Title" size="10" maxlength="10" value="<%=myHTMLEncode(Title)%>"></td>
					</tr>
					<tr>
						<td width="30%" bgcolor="#7DB1FF"><b>
              			<font size="1" face="Verdana"><%=getclientContactLngStr("LtxtPosition")%></font></b></td>
						<td height="18" class="GeneralTbl">
						<input type="text" name="Position" size="40" maxlength="90" value="<%=myHTMLEncode(Position)%>"></td>
					</tr>
					<tr>
						<td width="30%" bgcolor="#7DB1FF"><b>
              			<font size="1" face="Verdana"><%=getclientContactLngStr("DtxtAddress")%></font></b></td>
						<td height="18" class="GeneralTbl">
						<input type="text" name="Address" size="40" maxlength="90" value="<%=myHTMLEncode(Address)%>"></td>
					</tr>
					<tr>
						<td width="30%" bgcolor="#7DB1FF"><b>
              			<font size="1" face="Verdana"><%=getclientContactLngStr("DtxtPhone")%>&nbsp;1</font></b></td>
						<td height="18" class="GeneralTbl">
						<input type="text" name="Tel1" size="20" maxlength="20" value="<%=myHTMLEncode(Tel1)%>"></td>
					</tr>
					<tr>
						<td width="30%" bgcolor="#7DB1FF"><b>
              			<font size="1" face="Verdana"><%=getclientContactLngStr("DtxtPhone")%>&nbsp;2</font></b></td>
						<td height="18" class="GeneralTbl">
						<input type="text" name="Tel2" size="20" maxlength="20" value="<%=myHTMLEncode(Tel2)%>"></td>
					</tr>
					<tr>
						<td width="30%" bgcolor="#7DB1FF"><b>
              			<font size="1" face="Verdana"><%=getclientContactLngStr("LtxtMobile")%></font></b></td>
						<td height="18" class="GeneralTbl">
						<input type="text" name="Cellolar" size="20" maxlength="20" value="<%=myHTMLEncode(Cellolar)%>"></td>
					</tr>
					<tr>
						<td width="30%" bgcolor="#7DB1FF"><b>
              			<font size="1" face="Verdana"><%=getclientContactLngStr("DtxtFax")%></font></b></td>
						<td height="18" class="GeneralTbl">
						<input type="text" name="Fax" size="20" maxlength="20" value="<%=myHTMLEncode(Fax)%>"></td>
					</tr>
					<tr>
						<td width="30%" bgcolor="#7DB1FF"><b>
              			<font size="1" face="Verdana"><%=getclientContactLngStr("DtxtEMail")%></font></b></td>
						<td height="18" class="GeneralTbl">
						<input type="text" name="EMail" size="40" maxlength="100" value="<%=myHTMLEncode(EMail)%>"></td>
					</tr>
					<tr>
						<td width="30%" bgcolor="#7DB1FF"><b>
              			<font size="1" face="Verdana"><%=getclientContactLngStr("LtxtPager")%></font></b></td>
						<td height="18" class="GeneralTbl">
						<input type="text" name="Pager" size="30" maxlength="30" value="<%=myHTMLEncode(Pager)%>"></td>
					</tr>
					<tr>
						<td width="30%" bgcolor="#7DB1FF"><b>
              			<font size="1" face="Verdana"><%=getclientContactLngStr("DtxtObservations")%>&nbsp;1</font></b></td>
						<td height="18" class="GeneralTbl">
						<input type="text" name="Notes1" size="40" maxlength="100" value="<%=myHTMLEncode(Notes1)%>"></td>
					</tr>
					<tr>
						<td width="30%" bgcolor="#7DB1FF"><b>
              			<font size="1" face="Verdana"><%=getclientContactLngStr("DtxtObservations")%>&nbsp;2</font></b></td>
						<td height="18" class="GeneralTbl">
						<input type="text" name="Notes2" size="40" maxlength="100" value="<%=myHTMLEncode(Notes2)%>"></td>
					</tr>
					<tr>
						<td width="30%" bgcolor="#7DB1FF"><b>
              			<font size="1" face="Verdana"><%=getclientContactLngStr("DtxtPwd")%></font></b></td>
						<td height="18" class="GeneralTbl">
						<input type="text" name="Password" size="8" maxlength="8" value="<%=myHTMLEncode(Password)%>"></td>
					</tr>
					<% 
					Select Case myApp.LawsSet
						Case "MX", "CL", "CR", "GT", "US", "CA", "BR" %>
					<tr>
						<td width="30%" bgcolor="#7DB1FF"><b>
              			<font size="1" face="Verdana"><%=getclientContactLngStr("LtxtGender")%></font></b></td>
						<td height="18" class="GeneralTbl">
						<select size="1" name="Gender">
						<option value="">|L:txtSel|</option>
						<option value="M" <% If Gender = "M" Then %>selected<% End If %>>|D:txtMale|</option>
						<option value="F" <% If Gender = "F" Then %>selected<% End If %>>|D:txtFemale|</option>
						</select>
					</td>
					</tr>
					<tr>
						<td width="30%" bgcolor="#7DB1FF"><b>
              			<font size="1" face="Verdana"><%=getclientContactLngStr("LtxtProfesion")%></font></b></td>
						<td height="18" class="GeneralTbl">
						<input type="text" name="Profession" maxlength="50" size="40" value="<%=myHTMLEncode(Profession)%>"></td>
					</tr>
					<% End Select %>
					<tr>
						<td>&nbsp;</td>
						<td height="18" class="GeneralTbl">
						<input type="checkbox" name="SetDef" id="SetDef" <% If isDefault Then %>checked disabled<% End IF %> value="Y" style="background: background-image; border: 0px solid"><label for="SetDef"><%=getclientContactLngStr("LtxtSetAsDef")%></label></td>
					</tr>
					
					<% do while not rg.eof %>
					<tr>
						<td bgcolor="#7DB1FF" colspan="2"><b>
						<font size="1" face="Verdana"><% Select Case CInt(rg("GroupID"))
						Case -1 %><%=getclientContactLngStr("DtxtUDF")%><%
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
						<% If rcOpt("Query") = "Y" or rcOpt("TypeID") = "D" Then %></td><td width="16"><a href="#" onclick="javascript:document.frmCrd.U_<%=rcOpt("AliasID")%>.value = ''"><img border="0" src="../images/remove.gif" width="16" height="16"></a></td></tr></table><% End If %><% End If %></td>
					</tr>
                    <% 
                    rcOpt.movenext
                    loop 
                    rg.movenext
                    loop
                    rcOpt.Filter = "" %>
					<tr>
						<td colspan="2" class="style1">
						<input type="submit" value="<%=getclientContactLngStr("DtxtSave")%>" name="btnSave" onclick="return valFrm();"></td>
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
	<input type="hidden" name="cmd" value="contact">
	<input type="hidden" name="EditID" value='<%=Request("EditID")%>'>
	<input type="hidden" name="editVar" value="">
	<input type="hidden" name="returnCmd" value="newClientContact">
	</form>
    </table>
  </center>
</div>
<script type="text/javascript">
function getCal(AliasID)
{
	document.frmCrd.action = 'operaciones.asp';
	document.frmCrd.editVar.value = AliasID;
	document.frmCrd.cmd.value = 'UDFCal';
	document.frmCrd.submit();
}
function getVal(AliasID)
{
	document.frmCrd.action = 'operaciones.asp';
	document.frmCrd.editVar.value = AliasID;
	document.frmCrd.cmd.value = 'UDFQry';
	document.frmCrd.submit();
}
function valFrm()
{
	if (document.frmCrd.NewName.value == '')
	{
		alert('<%=getclientContactLngStr("LtxtValNam")%>');
		document.frmCrd.NewName.focus();
		return false;
	}
	<%
	If rcOpt.recordcount > 0 Then rcOpt.movefirst
	do while not rcOpt.eof
	If rcOpt("NullField") = "Y" or rcOpt("TypeID") = "N" or rcOpt("TypeID") = "B" Then
		If rcOpt("NullField") = "Y" Then %>
		if (document.frmCrd.U_<%=rcOpt("AliasID")%>.value == '') 
		{
			alert('<%=getclientContactLngStr("LtxtValFld")%>'.replace('{0}', '<%=Replace(rcOpt("Descr"), "'", "\'")%>'));
			document.frmCrd.U_<%=rcOpt("AliasID")%>.focus
			return false; 
		}
		<% End If
		If rcOpt("TypeID") = "B" or rcOpt("TypeID") = "N" Then %>
		if (document.frmCrd.U_<%=rcOpt("AliasID")%>.value != '') 
		{
			if (!MyIsNumeric(document.frmCrd.U_<%=rcOpt("AliasID")%>.value)) 
			{
				alert('<%=getclientContactLngStr("DtxtValNumVal")%>');
				document.frmCrd.U_<%=rcOpt("AliasID")%>.focus
				return false; 
			}
		}
		<% End If
	End If
	rcOpt.movenext
	loop
	rcOpt.close  %>
	return true;
}
</script>
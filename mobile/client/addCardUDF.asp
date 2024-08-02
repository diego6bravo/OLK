<% addLngPathStr = "client/" %><!--#include file="lang/addCardUDF.asp" -->
<%


set rg = Server.CreateObject("ADODB.RecordSet")
sql = "select T0.GroupID, IsNull(T1.AlterGroupName, T0.GroupName) GroupName " & _
"from OLKCUFDGroups T0 " & _  
"left outer join OLKCUFDGroupsAlterNames T1 on T1.TableID = T0.TableID and T1.GroupID = T0.GroupID and T1.LanID = " & Session("LanID") & " " & _
"where T0.TableID = 'OCRD' and exists(select '' from CUFD X0 left outer join OLKCUFD X1 on X1.TableID = X0.TableID and X1.FieldID = X0.FieldID where X0.TableID = T0.TableID and IsNull(X1.GroupID, -1) = T0.GroupID and AType in ('" & userType & "', 'T') and OP in ('T','P') and Active = 'Y')  " & _  
"order by T0.[Order]"
set rg = conn.execute(sql)
If Request("GroupID") <> "" Then 
	GroupID = Request("GroupID") 
Else 
	If Not rg.Eof Then GroupID = rg("GroupID") Else GroupID = -1
End If

sqlAddStr = ""
sql = "select (select SDKID collate database_default from r3_obscommon..tcif where companydb = N'" & Session("OlkDB") & "')++AliasID As InsertID, T0.TypeID, RTable " & _
	  "from cufd T0 left outer join OLKCUFD T1 on T1.TableID = T0.TableID and T1.FieldID = T0.FieldID " & _
	  "where T0.TableId = 'OCRD' and AType in ('V', 'T') and OP in ('T','P') and Active = 'Y' and IsNull(T1.GroupID, -1) = " & GroupID
set rd = conn.execute(sql)
do while not rd.eof
	sqlAddStr = sqlAddStr & ", " & rd("InsertID")
rd.movenext
loop



sql = "select CardCode, IsNull(CardName, '') CardName" & sqlAddStr & ", " & _
"(select Command from R3_ObsCommon..TLOG where LogNum = T0.LogNum) Command " & _
"from R3_ObsCommon..TCRD T0 " & _
"where T0.LogNum = " & Session("CrdRetVal")
set rs = conn.execute(sql)
isUpdate = rs("Command") = "U"
CardCode = rs("CardCode")
CardName = rs("CardName")
 %>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
  <form method="post" action="client/submitClient.asp" name="frmCrd" onsubmit="return valFrm();">
    <tr>
      <td>
      <img src="images/spacer.gif" width="100%" height="1" border="0" alt></td>
    </tr>
    <tr>
      <td bgcolor="#9BC4FF">
      <table border="0" cellpadding="0" bordercolor="#111111" width="100%">
        <tr>
          <td width="100%">
          <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><% If Not isUpdate Then %><%=getaddCardUDFLngStr("LttlNewClient")%><% Else %><%=getaddCardUDFLngStr("LttlEditClient")%><% End If %> - <%=getaddCardUDFLngStr("DtxtUDF")%>
          </font></b></td>
        </tr>
        <tr>
          <td width="100%">
         <!--#include file="clientMenu.asp"--></td>
        </tr>
        <tr>
          <td width="100%">
            <table border="0" cellpadding="0"  bordercolor="#111111" width="100%">
            <tr>
              <td width="30%" bgcolor="#7DB1FF"><b>
              <font size="1" face="Verdana"><%=getaddCardUDFLngStr("DtxtCode")%></font></b></td>
              <td width="70%" bgcolor="#8CBAFF"><p>
              <font size="1" face="Verdana"><%=myHTMLEncode(CardCode)%></font></td>
            </tr>
			<tr>
              <td width="30%" bgcolor="#7DB1FF"><b>
              <font size="1" face="Verdana"><%=getaddCardUDFLngStr("DtxtName")%></font></b></td>
              <td width="70%" bgcolor="#8CBAFF"><p>
              <font size="1" face="Verdana"><%=myHTMLEncode(CardName)%></font></td>
            </tr>
			<tr>
              <td width="30%" bgcolor="#7DB1FF"><b>
              <font size="1" face="Verdana"><%=getaddCardUDFLngStr("DtxtGroup")%></font></b></td>
              <td width="70%" bgcolor="#8CBAFF"><p>
              <select name="newGroupID" size="1" onchange="if(valFrm()){document.frmCrd.changeGroup.value='Y';submit();}else document.frmCrd.newGroupID.value = document.frmCrd.GroupID.value;">
				<% do while not rg.eof %><option <% If CInt(GroupID) = CInt(rg("GroupID")) Then %>selected<% End If %> value="<%=rg("GroupID")%>"><%=rg("GroupName")%></option><% rg.movenext
				loop %>
				</select>
				<input type="hidden" name="changeGroup" value="N"></td>
            </tr>
            <%
            set rcOpt = Server.CreateObject("ADODB.RecordSet")
			sql = "select T0.FieldID, AliasID, IsNull(alterDescr, Descr) Descr, TypeID, SizeID, Dflt, NotNull, Pos, RTable, " & _
				  "Case When Exists(select 'A' from UFD1 where TableId = T0.TableId and FieldId = T0.FieldId) " & _
				  "Then 'Y' Else 'N' End As DropDown, NullField, Query, " & _
				  "(select SDKID collate database_default from r3_obscommon..tcif where companydb = '" & Session("OlkDB") & "')++AliasID As InsertID " & _
				  "from cufd T0 " & _
				  "left outer join OLKCUFD T1 on T1.TableID = T0.TableID and T1.FieldID = T0.FieldID " & _
				  "left outer join OLKCUFDAlterNames T2 on T2.TableID = T0.TableID and T2.FieldID = T0.FieldID and T2.LanID = " & Session("LanID") & " " & _
				  "where T0.TableId = 'OCRD' and AType in ('" & userType & "', 'T') and OP in ('T','P') and Active = 'Y' and IsNull(T1.GroupID, -1) = " & GroupID
			rcOpt.open sql, conn, 3, 1
						
			do while not rcOpt.eof 
            AliasID = rcOpt("InsertID")
            If Request.Form.Count = 0 Then
	            fldVal = rs(AliasID)
	            If rcOpt("TypeID") = "D" Then fldVal = FormatDate(fldVal, False)
	        Else
	        	fldVal = Request("U_" & rcOpt("AliasID"))
	        End If %>
            <tr>
              <td width="33%" bgcolor="#7DB1FF"><b>
                      <font size="1" face="Verdana">&nbsp;<%=rcOpt("Descr")%><% If rcOpt("NullField") = "Y" Then %><font color="red">*</font><% End If %></font></b></td>
              <td width="67%" colspan="2">
        		<% If rcOpt("DropDown") = "Y" or Not IsNull(rcOpt("RTable")) then 
        		If rcOpt("DropDown") = "Y" Then
	        		sql = "select FldValue, IsNull(AlterDescr, Descr) Descr " & _
									"from UFD1 T0 " & _
									"left outer join OLKUFD1AlterNames T1 on T1.TableID = T0.TableID and T1.FieldID = T0.FieldID and T1.IndexID = T0.IndexID and T1.LanID = " & Session("LanID") & " " & _
									"where T0.tableid = 'OCRD' and T0.FieldId = " & rcOpt("FieldId")
				Else
					sql = "select Code FldValue, Name Descr from [@" & rcOpt("RTable") & "] order by 2"
				End If
				set rctn = conn.execute(sql) %>
				<font color="#4783C5">
				<select size="1" name="U_<%=rcOpt("AliasID")%>" class="input" style="font-size:10px; width:100%; font-family:Verdana">
				<option></option>
				<% do while not rctn.eof %>
				<option value="<%=rctn("FldValue")%>" <% If fldVal = rctn("FldValue") Then %>selected<% ElseIf rctn("FldValue") = rcOpt("Dflt") and IsNull(fldVal) Then %>selected<% End If %>><%=rctn("Descr")%></option>
				<% rctn.movenext
				loop
				rctn.close %></select></font><font size="1" color="#4783C5">
				<% Else %>
				<% If rcOpt("Query") = "Y" or rcOpt("TypeID") = "D" Then %><table width="100%" cellspacing="0" cellpadding="0"><tr><td width="16"><img border="0" src="<% If rcOpt("Query") = "Y" Then %>../images/flechaselec2.gif<% Else %>images/cal.gif<% End If %>" <% If rcOpt("TypeID") = "D" Then %>onclick="javascript:getCal('<%=rcOpt("AliasID")%>')"<% End If %> <% If rcOpt("Query") = "Y" Then %>onclick="javascript:getVal('<%=rcOpt("AliasID")%>')"<% End If %>></td><td><% End If %>
				<input <% If rcOpt("Query") = "Y" or rcOpt("TypeID") = "D" Then %>readonly<% End If %> type="text" name="U_<%=rcOpt("AliasID")%>" size="<% If rcOpt("TypeID") = "A" Then %>43<% Else %>12<% End If %>" class="input" value="<% If fldVal <> "" Then %><%=fldVal%><% Else %><%=rcOpt("Dflt")%><% End If %>" onclick="this.selectionStart=0;this.selectionEnd=this.value.length;" style="width: 100%; font-family: Verdana; font-size: 10px">
				<% If rcOpt("Query") = "Y" or rcOpt("TypeID") = "D" Then %></td><td width="16"><img border="0" src="../images/remove.gif" width="16" height="16" onclick="javascript:document.frmCrd.U_<%=rcOpt("AliasID")%>.value = ''"></td></tr></table><% End If %><% End If %></td>
            </tr>
                    <% 
                    rcOpt.movenext
                    loop 
                    If rcOpt.RecordCount > 0 then rcOpt.movefirst %>
          </table>
           </td>
        </tr>
        <tr>
          <td width="100%">&nbsp;</td>
        </tr>
      </table>
      </td>
    </tr>
	<input type="hidden" name="cmd" value="UDF">
	<input type="hidden" name="editVar" value="">
	<input type="hidden" name="returnCmd" value="newClientUDF">
	<input type="hidden" name="GroupID" value="<%=GroupID%>">
	</form>
    </table>
  </center>
</div>
<script language="javascript">
function valFrm() 
{
	<% 
	do while not rcOpt.eof
	If rcOpt("NotNull") = "Y" or rcOpt("NullField") = "Y" Then %>
	if (document.frmCrd.U_<%=rcOpt("AliasID")%>.value == '') {
		alert('<%=getaddCardUDFLngStr("LtxtValFld")%>'.replace('{0}', '<%=Replace(rcOpt("Descr"), "'", "\'")%>'));
		document.frmCrd.U_<%=rcOpt("AliasID")%>.focus
		return false; }
	<% End If
	If rcOpt("TypeID") = "B" or rcOpt("TypeID") = "N" Then %>
	if (document.frmCrd.U_<%=rcOpt("AliasID")%>.value != '') 
	{
		if (!IsNumeric(document.frmCrd.U_<%=rcOpt("AliasID")%>.value)) 
		{
			alert('<%=getaddCardUDFLngStr("DtxtValNumVal")%>');
			document.frmCrd.U_<%=rcOpt("AliasID")%>.focus
			return false; 
		}
	}
	<% End If
	rcOpt.movenext
	loop %>
	document.frmCrd.cmd.value = 'UDF';
	document.frmCrd.action = 'client/submitClient.asp';
	return true;
}

function IsNumeric(sText)
{
   var ValidChars = "0123456789.";
   var IsNumber=true;
   var Char;

 
   for (i = 0; i < sText.length && IsNumber == true; i++) 
      { 
      Char = sText.charAt(i); 
      if (ValidChars.indexOf(Char) == -1) 
         {
         IsNumber = false;
         }
      }
   return IsNumber;
   
}
   
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
</script>

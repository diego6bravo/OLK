<% addLngPathStr = "activity/" %>
<!--#include file="lang/activityUDF.asp" -->

<head>
<style type="text/css">
.style1 {
				font-family: Verdana;
				font-size: xx-small;
}
.style2 {
				font-family: Verdana;
}
.style3 {
				font-size: xx-small;
}
.style4 {
				background-color: #75ACFF;
}
.style5 {
				font-family: Verdana;
				font-size: xx-small;
				background-color: #75ACFF;
}
</style>
</head>

<% 
set rc = server.createobject("ADODB.RecordSet")
set rd = server.createobject("ADODB.RecordSet")
set rctn = server.createobject("ADODB.RecordSet")
ReadOnly = Session("ActReadOnly")

If Not Session("ActReadOnly") Then
	SDKID = "(select SDKID collate database_default from r3_obscommon..tcif where companydb = '" & Session("OlkDB") & "')"
Else
	SDKID = "'U_'"
End If

set rg = Server.CreateObject("ADODB.RecordSet")
sql = "select T0.GroupID, IsNull(T1.AlterGroupName, T0.GroupName) GroupName " & _
"from OLKCUFDGroups T0 " & _  
"left outer join OLKCUFDGroupsAlterNames T1 on T1.TableID = T0.TableID and T1.GroupID = T0.GroupID and T1.LanID = " & Session("LanID") & " " & _
"where T0.TableID = 'OCLG' and exists(select '' from CUFD X0 left outer join OLKCUFD X1 on X1.TableID = X0.TableID and X1.FieldID = X0.FieldID where X0.TableID = T0.TableID and IsNull(X1.GroupID, -1) = T0.GroupID and AType in ('" & userType & "', 'T') and OP in ('T','P') and Active = 'Y')  " & _  
"order by T0.[Order]"
set rg = conn.execute(sql)
If Request("GroupID") <> "" Then 
	GroupID = Request("GroupID") 
Else 
	If Not rg.Eof Then GroupID = rg("GroupID") Else GroupID = -1
End If

set rcOpt = Server.CreateObject("ADODB.RecordSet")
sql = "select T0.FieldID, AliasID, IsNull(alterDescr, Descr) Descr, TypeID, SizeID, Dflt, NotNull, Pos, RTable, " & _
	"Case When Exists(select 'A' from UFD1 where TableId = T0.TableId and FieldId = T0.FieldId) " & _
	"Then 'Y' Else 'N' End As DropDown, NullField, Query, " & _
	SDKID & "++AliasID As InsertID, T0.EditType " & _
	"from cufd T0 " & _
	"left outer join OLKCUFD T1 on T1.TableID = T0.TableID and T1.FieldID = T0.FieldID " & _
	"left outer join OLKCUFDAlterNames T2 on T2.TableID = T0.TableID and T2.FieldID = T0.FieldID and T2.LanID = " & Session("LanID") & " " & _
	"where T0.TableId = 'OCLG' and IsNull(T1.GroupID, -1) = " & GroupID & " and AType in ('" & userType & "', 'T') and OP in ('T','P') and Active = 'Y' " & _
	"order by IsNull(Pos, 'D'), IsNull(T1.[Order],32727) "
rcOpt.open sql, conn, 3, 1

sql = "select AliasID, NullField, Descr, TypeID " & _
	"from cufd T0 " & _
	"left outer join OLKCUFD T1 on T1.TableID = T0.TableID and T1.FieldID = T0.FieldID " & _
	"where T0.TableId = 'OCLG' and IsNull(T1.GroupID, -1) = " & GroupID & " and AType in ('" & userType & "', 'T') and OP in ('T','P') and Active = 'Y' and (NullField = 'Y' or TypeID in ('N', 'B')) " & _
	"Order By IsNull(Pos, 'D'), IsNull(T1.[Order],32727)" 
rctn.open sql, conn, 3, 1

If rctn.recordcount > 0 Then chkOpt = True

addCols = ""

If rcOpt.RecordCount > 0 Then
	set rcOptVals = Server.CreateObject("ADODB.RecordSet")
	sql = "select " & SDKID & "++AliasID As InsertID, TypeID " & _
		"from cufd T0 " & _
		"left outer join OLKCUFD T1 on T1.TableID = T0.TableID and T1.FieldID = T0.FieldID " & _
		"where T0.TableId = 'OCLG' and AType in ('" & userType & "','T') and OP in ('T','P')  and Active = 'Y'"
	rcOptVals.open sql, conn, 3, 1
	
	do while not rcOptVals.eof
		addCols = addCols & ", "
		addCols = addCols & rcOptVals("InsertID")
	rcOptVals.movenext
	loop
End If

If Not Session("ActReadOnly") Then
	sql = "select T0.ClgCode, T0.CardCode" & addCols  & " " & _
	"from R3_ObsCommon..TCLG T0 " & _
	"where T0.LogNum = " & Session("ActRetVal")
Else 
	sql = "select T0.ClgCode, T0.CardCode" & addCols  & " " & _
	"from OCLG T0 " & _
	"where T0.ClgCode = " & Session("ActRetVal")
End If
set rs = conn.execute(sql)
ClgCode = rs("ClgCode")
 %>
 <script type="text/javascript">
function valFrm() {
	<%
	If Not ReadOnly Then
	If rctn.recordcount > 0 Then rctn.movefirst
	do while not rctn.eof
	If rctn("NullField") = "Y" Then %>
	if (document.frmUDF.U_<%=rctn("AliasID")%>.value == '') 
	{
		alert('<%=getactivityUDFLngStr("LtxtValFld")%>'.replace('{0}', '<%=Replace(rctn("Descr"), "'", "\'")%>'));
		document.frmUDF.U_<%=rctn("AliasID")%>.focus
		return false; 
	}
	<% End If
	If rctn("TypeID") = "B" or rctn("TypeID") = "N" Then %>
	if (document.frmUDF.U_<%=rctn("AliasID")%>.value != '') 
	{
		if (!IsNumeric(document.frmUDF.U_<%=rctn("AliasID")%>.value)) 
		{
			alert('<%=getactivityUDFLngStr("DtxtValNumVal")%>');
			document.frmUDF.U_<%=rctn("AliasID")%>.focus
			return false; 
		}
	}
	<% End If
	rctn.movenext
	loop
	rctn.close
	End If  %>
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
</script>

<div align="center">
				<center>
				<table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111" bgcolor="#9BC4FF">
								<form name="frmUDF" method="post" action="activity/actSubmit.asp" onsubmit="return valFrm();">
												<input type="hidden" name="cmd" value="udf">
								      			<input type="hidden" name="editVar" value="">
								      			<input type="hidden" name="returnCmd" value="activityUDF">
								      			<input type="hidden" name="GroupID" value="<%=GroupID%>">
												<tr>
          <td width="100%" align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
          <table cellpadding="0" border="0">
						<tr>
										<td><img src="images/icon_activity_<% If Not IsNull(ClgCode) Then %>S<% Else %>O<% End If %>.gif"></td>
										<td><b><font face="Verdana" size="1"><%=getactivityUDFLngStr("DtxtActivity")%>&nbsp;#<% If Not IsNull(ClgCode) Then Response.Write ClgCode Else Response.Write Session("ActRetVal") %>&nbsp;-&nbsp;<%=getactivityUDFLngStr("DtxtUDF")%></font></b></td>
						</tr>
			</table>
          </td>
												</tr><tr>
												          <td width="100%">
												          <!--#include file="activityMenu.asp"--></td>
												        </tr>
												<tr>
																<td><b>
																<font face="Verdana" size="1">
																<%=getactivityUDFLngStr("DtxtGroup")%>:&nbsp;</font></b><select name="newGroupID" size="1" onchange="if(valFrm()){document.frmUDF.changeGroup.value='Y';submit();}else document.frmUDF.newGroupID.value = document.frmUDF.GroupID.value;">
																<% do while not rg.eof %><option <% If CInt(GroupID) = CInt(rg("GroupID")) Then %>selected<% End If %> value="<%=rg("GroupID")%>"><%=rg("GroupName")%></option><% rg.movenext
																loop %>
																</select>
								      							<input type="hidden" name="changeGroup" value="N">
																</td>
												</tr>
        <tr>
          <td width="100%">
          <table border="0" cellpadding="0" bordercolor="#111111" width="100%">

					<% do while not rcOpt.eof 
                    AliasID = rcOpt("InsertID")
                    If Request.Form.Count = 0 Then
	                    fldVal = rs(AliasID)
	                    If rcOpt("TypeID") = "D" Then fldVal = FormatDate(fldVal, False)
	                Else
	                	fldVal = Request("U_" & rcOpt("AliasID"))
	                End If %>
            <tr>
              <td><b><font size="1" face="Verdana">&nbsp;<%=rcOpt("Descr")%><% If rcOpt("NullField") = "Y" Then %><font color="red">*</font><% End If %></font></b></td>
			</tr>
			<tr>
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
				set rctn = conn.execute(sql) %>
				<font color="#4783C5">
				<select <% If ReadOnly Then %>disabled<% End If %> size="1" name="U_<%=rcOpt("AliasID")%>" class="input" style="font-size:10px; width:100%; font-family:Verdana">
				<option></option>
				<% do while not rctn.eof %>
				<option value="<%=rctn("FldValue")%>" <% If fldVal = rctn("FldValue") Then %>selected<% ElseIf rctn("FldValue") = rcOpt("Dflt") and IsNull(fldVal) Then %>selected<% End If %>><%=rctn("Descr")%></option>
				<% rctn.movenext
				loop %></select></font><font size="1" color="#4783C5">
				<% ElseIf rcOpt("EditType") = "I" Then
				If fldVal <> "" Then img = fldVal Else img = "n_a.gif" %>
				<table cellpadding="2" cellspacing="2" border="0">
					<tr>
						<td align="center" colspan="2"><img id="img<%=rcOpt("AliasID")%>" src="pic.aspx?filename=<%=img%>&dbName=<%=Session("olkdb")%>">
						<input type="hidden" name="U_<%=rcOpt("AliasID")%>" id="U_<%=rcOpt("AliasID")%>" value="<%=fldVal%>"></td>
					</tr>
					<tr>
						<td align="center"><input type="button" name="btnRem" value="<%=getactivityUDFLngStr("DtxtClear")%>" onclick="document.getElementById('U_<%=rcOpt("AliasID")%>').value = '';document.getElementById('img<%=rcOpt("AliasID")%>').src='pic.aspx?filename=n_a.gif&dbName=<%=Session("olkdb")%>';"></td>
						<td align="center"><input type="button" name="btnUpload" value="<%=getactivityUDFLngStr("DtxtChange")%>"></td>
					</tr>
				</table>
				<% Else %>
				<% If (rcOpt("Query") = "Y" or rcOpt("TypeID") = "D") and not ReadOnly Then %><table width="100%" cellspacing="0" cellpadding="0"><tr><td width="16"><a href="#" <% If rcOpt("TypeID") = "D" Then %>onclick="javascript:getCalUDF('<%=rcOpt("AliasID")%>')"<% End If %> <% If rcOpt("Query") = "Y" Then %>onclick="javascript:getValUDF('<%=rcOpt("AliasID")%>')"<% End If %>><img border="0" src="<% If rcOpt("Query") = "Y" Then %>../images/<%=Session("rtl")%>flechaselec2.gif<% Else %>images/cal.gif<% End If %>"></a></td><td><% End If %>
				<input <% If rcOpt("Query") = "Y" or rcOpt("TypeID") = "D" or ReadOnly Then %>readonly<% End If %> type="text" name="U_<%=rcOpt("AliasID")%>" size="<% If rcOpt("TypeID") = "A" Then %>43<% Else %>12<% End If %>" class="input" value="<% If fldVal <> "" Then %><%=fldVal%><% Else %><%=rcOpt("Dflt")%><% End If %>" onclick="this.selectionStart=0;this.selectionEnd=this.value.length;" style="width: 100%; font-family: Verdana; font-size: 10px" <% If rcOpt("TypeID") = "D" Then %>onclick="javascript:getCalUDF('<%=rcOpt("AliasID")%>')"<% End If %> <% If rcOpt("Query") = "Y" Then %>onclick="javascript:getValUDF('<%=rcOpt("AliasID")%>')"<% End If %>>
				<% If (rcOpt("Query") = "Y" or rcOpt("TypeID") = "D") and not ReadOnly Then %></td><td width="16"><a href="#" onclick="javascript:document.frmUDF.U_<%=rcOpt("AliasID")%>.value = ''"><img border="0" src="../images/remove.gif" width="16" height="16"></a></td></tr></table><% End If %><% End If %></td>
            </tr>
                    <% 
                    rcOpt.movenext
                    loop 
                    If rcOpt.RecordCount > 0 then rcOpt.movefirst %>
		</table>
		</td>
	</tr>
								<!--#include file="activityBottom.asp"-->

					
			</form>
				</table>
				</center></div>

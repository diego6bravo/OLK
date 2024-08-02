<!--#include file="lang/adSearchProp.asp" -->
<link rel="stylesheet" href="Reportes/style.css">
<% 

set rs = Server.CreateObject("ADODB.RecordSet")

sql = 	"select IsNull(T2.alterName, T0.Name) Name " & _
		"from OLKCustomSearch T0 " & _
		"left outer join OLKCustomSearchAlterNames T2 on T2.ObjectCode = T0.ObjectCode and T2.ID = T0.ID and T2.LanID = " & Session("LanID") & " " & _
		"where T0.ObjectCode = " & Request("adObjID") & " and T0.ID = " & Request.Form("ID")
		
set rs = conn.execute(sql)
set rd = Server.CreateObject("ADODB.RecordSet")
rsName = rs(0)


Select Case CInt(Request("adObjID"))
	Case 4
		sql = 	"select T0.ItmsTypCod, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITG', 'ItmsGrpNam', T0.ItmsTypCod, T0.ItmsGrpNam) ItmsGrpNam " & _
				"from OITG T0 " & _
				"inner join OLKCustomSearchProp T1 on T1.ObjectCode = 4 and T1.ID = " & Request.Form("ID") & " and T1.PropID = T0.ItmsTypCod " & _
				"where T1.Active = 'Y' " & _
				"order by T1.Ordr"
	Case 2
		sql = 	"select T0.GroupCode, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(1, 'OCQG', 'GroupName', GroupCode, GroupName) GroupName " & _
				"from OCQG T0 " & _
				"inner join OLKCustomSearchProp T1 on T1.ObjectCode = 2 and T1.ID = " & Request.Form("ID") & " and T1.PropID = T0.GroupCode " & _
				"where T1.Active = 'Y' " & _
				"order by T1.Ordr"
End Select

set rs = conn.execute(sql)

	
selVals = Request("chkQryGroup")
If selVals <> "" Then
	arrVals = Split(selVals, ", ")
End If

%>
<table border="0" cellspacing="0" width="100%" id="table1">
	<tr class="TblTltMnu">
		<td colspan="2"><img border="0" src="images/arrow_menu.gif" width="9" height="6">&nbsp;<%=rsName%> - <%=getadSearchPropLngStr("DtxtProp")%></td>
	</tr>
	<form method="POST" name="frmViewRep" action="operaciones.asp">
	<tr class="TblAfueraMnu">
		<td colspan="2" align="center">
		<table border="0" cellpadding="0" width="100%" id="table2">
			<% i = 0
			do while not rs.eof
			isChecked = False
			If selVals <> "" Then
				For j = 0 to UBound(arrVals)
					If CInt(arrVals(j)) = CInt(rs(0)) Then
						isChecked = True
						Exit For
					End If
				Next
			End If %>
			<tr class="TblAfueraMnu">
				<td><input type="checkbox" id="chk<%=i%>" <% If isChecked Then %>checked<% End If %> name="chk" value="<%=rs(0)%>"><label id="txt<%=i%>" for="chk<%=i%>"><%=rs(1)%></label></td>
			</tr>
		  <% i = i + 1
		  rs.movenext
		  loop %>
		</table>
		</td>
	</tr>
	<tr class="TblAfueraMnu">
		<td colspan="2">
		<p align="center">
		</td>
	</tr>
		<% 	For each itm in Request.Form %>
		<input type="hidden" name="<%=itm%>" value="<%=Request(itm)%>">
		<% 	Next %>
	<tr class="TblAfueraMnu">
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	</tr>
	<tr class="TblAfueraMnu">
		<td><input type="submit" name="btnAccept" value="<%=getadSearchPropLngStr("DtxtAccept")%>" onclick="javascript:doAccept();"></td>
		<td>
		<p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
		<input type="submit" name="btnCancel" value="<%=getadSearchPropLngStr("DtxtCancel")%>" onclick="javascript:goBackToVars();"></td>
	</tr>
	</form>
</table>
<script type="text/javascript">
function doAccept()
{
	if (document.frmViewRep.chk)
	{
		var chk = document.frmViewRep.chk;
		
		var retVal = '';
		var retValDesc = '';
		
		if (chk.length)
		{
			for (var i = 0;i<chk.length;i++)
			{
				if (chk[i].checked)
				{
					if (retVal != '')
					{
						retVal += ', ';
						retValDesc += '\n';
					}
					
					retVal += chk[i].value;
					retValDesc += document.getElementById(chk[i].id.replace('chk', 'txt')).innerText;
				}
			}
		}
		else
		{
			if (chk.checked)
			{
				retVal = chk.value;
				retValDesc = document.getElementById(chk.id.replace('chk', 'txt')).innerText;
			}
		}
		document.frmViewRep.chkQryGroup.value=retVal;
		document.frmViewRep.chkQryGroupDesc.value=retValDesc;
	}
	goBackToVars();
}

function goBackToVars()
{
	document.frmViewRep.isSubmit.value = 'R';
	document.frmViewRep.cmd.value='adSearch';
}
</script>
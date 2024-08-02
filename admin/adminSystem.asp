<!--#include file="top.asp" -->
<!--#include file="lang/adminSystem.asp" -->

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style type="text/css">
.style1 {
	font-family: Verdana;
	font-size: xx-small;
	color: #31659C;
}
</style>
</head>
<% If Request("Restore") = "OK" Then %>
<script language="javascript">
alert('<%=getadminSystemLngStr("LtxtRestoreOK")%>');
</script>
<% End If %>
<% 
If Request.Form.Count > 0 Then
	
	If Request("AllowSavePwd") = "Y" Then AllowSavePwd = "Y" Else AllowSavePwd = "N"
	If Request("ShowDbName") = "Y" Then ShowDbName = "Y" Else ShowDbName = "N"
	If Request("SingleSignOn") = "Y" Then SingleSignOn = "Y" Else SingleSignOn = "N"
	
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "OLKSaveSystemSettings"
	cmd.Parameters.Refresh()
	cmd("@AllowSavePwd") = AllowSavePwd
	cmd("@ShowDbName") = ShowDbName
	cmd("@SingleSignOn") = SingleSignOn
	cmd.execute()
	
	Application("AllowSavePwd") = AllowSavePwd = "Y"
	Application("ShowDbName") = ShowDbName = "Y"
	Application("SingleSignOn") = SingleSignOn = "Y"
End If 
 %>
<table border="0" cellpadding="0" width="98%">
<form name="frm" method="post" action="adminSystem.asp">
	<tr>
		<td height="15"></td>
	</tr>
	<tr>
		<td bgcolor="#E1F3FD"><b><font color="#31659C" size="1" face="Verdana">&nbsp;<%=getadminSystemLngStr("LttlSysProp")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif">
		<font face="Verdana" color="#4783C5" size="1"><%=getadminSystemLngStr("LttlSysPropNote")%></font></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr>
				<td class="style1">
				<img src="images/ganchito.gif"><font face="Verdana" size="1"> </font>
				<input type="checkbox" name="AllowSavePwd" value="Y" <% If myApp.AllowSavePwd Then %>checked<% End If %> id="AllowSavePwd" class="noborder" style="width: 20px"><font color="#4783C5" face="Verdana" size="1"><label for="AllowSavePwd"><%=getadminSystemLngStr("LtxtSaveInPwd")%></label></font></td>
				<td class="style2">
				&nbsp;</td>
			</tr>
			<tr>
				<td class="style1">
				<img src="images/ganchito.gif"><font face="Verdana" size="1"> </font>
				<input type="checkbox" name="ShowDbName" value="Y" <% If myApp.ShowDbName Then %>checked<% End If %> id="ShowDbName" class="noborder" style="width: 20px"><font color="#4783C5" face="Verdana" size="1"><label for="ShowDbName"><%=getadminSystemLngStr("LtxtShowDBName")%></label></font></td>
				<td class="style2">
				&nbsp;</td>
			</tr>
			<tr>
				<td class="style1">
				<img src="images/ganchito.gif"><font face="Verdana" size="1"> </font>
				<input type="checkbox" name="SingleSignOn" value="Y" <% If myApp.SingleSignOn Then %>checked<% End If %> id="SingleSignOn" class="noborder" style="width: 20px"><font color="#4783C5" face="Verdana" size="1"><label for="SingleSignOn"><%=getadminSystemLngStr("LtxtSingleSignOn")%></label></font></td>
				<td class="style2">
				&nbsp;</td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table5">
			<tr>
				<td width="77">
				<input type="submit" value="<%=getadminSystemLngStr("DtxtSave")%>" name="B1" class="OlkBtn"></td>
				<td><hr color="#0D85C6" size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
</form>
</table>
<% If Session("olkdb") <> "" Then %>
<table border="0" cellpadding="0" width="98%">
	<form name="frmTest">
	<tr>
		<td height="15"></td>
	</tr>
	<tr>
		<td bgcolor="#E1F3FD"><b><font color="#31659C" size="1" face="Verdana">&nbsp;<%=getadminSystemLngStr("LtxtTest")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif">
		<font face="Verdana" color="#4783C5" size="1"><%=getadminSystemLngStr("LttlTestNote")%></font></td>
	</tr>
	<tr>
		<td>
		<div style="width: 100%; height: 400px; overflow: auto;" id="divTest">
		<div>
		<table cellpadding="0" width="100%" border="0">
			<tr class="TblRepTltSub" style="height: 16px; text-align:center;">
				<td width="10"></td>
				<td><%=getadminSystemLngStr("DtxtType")%></td>
				<td><%=getadminSystemLngStr("DtxtName")%></td>
				<td><%=getadminSystemLngStr("LtxtLastCheck")%></td>
				<td style="width: 100px;"><%=getadminSystemLngStr("DtxtValid")%></td>
			</tr>
			<%
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKGetTest" & Session("ID")
			cmd.Parameters.Refresh()
			set rs = Server.CreateObject("ADODB.RecordSet")
			set rs = cmd.execute()
			
			set rd = Server.CreateObject("ADODB.RecordSet")
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKGetTestRepData" & Session("ID")
			cmd.Parameters.Refresh()
			rd.open cmd, , 3, 1
			
			set rf = Server.CreateObject("ADODB.RecordSet")
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKGetTestFlowData" & Session("ID")
			cmd.Parameters.Refresh()
			rf.open cmd, , 3, 1

			
			do while not rs.eof
			Select Case rs("Type")
				Case 0
					typeDesc = getadminSystemLngStr("LtxtTaskMon")
					typeLink = "adminInformerEdit.asp?ID={0}"
				Case 1
					typeDesc = getadminSystemLngStr("LtxtPrintTtl")
					typeLink = "adminPrintTitle.asp?edit=Y&rI={0}"
				Case 2
					typeDesc = getadminSystemLngStr("LtxtCardDetails")
					typeLink = "adminCardOpt.asp?edit=Y&rI={0}"
				Case 3
					typeDesc = getadminSystemLngStr("LtxtItemDetails")
					typeLink = "adminInvOpt.asp?edit=Y&rI={0}"
				Case 4
					typeDesc = getadminSystemLngStr("LtxtBatchDetails")
					typeLink = "adminBatchOpt.asp?edit=Y&rI={0}"
				Case 5
					typeDesc = getadminSystemLngStr("DtxtUDF") & " - "
					Select Case rs("TypeID")
						Case "CRD1"
							typeDesc = typeDesc & getAdminSystemLngStr("DtxtClients") & " - " & getAdminSystemLngStr("DtxtAddress")
						Case "OCLG"
							typeDesc = typeDesc & getAdminSystemLngStr("DtxtActivities")
						Case "OCPR"
							typeDesc = typeDesc & getAdminSystemLngStr("DtxtClients") & " - " & getAdminSystemLngStr("DtxtContact")
						Case "OCRD"
							typeDesc = typeDesc & getAdminSystemLngStr("DtxtClients")
						Case "OINV"
							typeDesc = typeDesc & getAdminSystemLngStr("DtxtComDocs")
						Case "INV1"
							typeDesc = typeDesc & getAdminSystemLngStr("DtxtComDocs") & " - " & getAdminSystemLngStr("DtxtLines")
						Case "OITM"
							typeDesc = typeDesc & getAdminSystemLngStr("DtxtItems")
						Case "ORCT"
							typeDesc = typeDesc & getAdminSystemLngStr("DtxtReceipts")
					End Select
					typeLink = "adminCUFD.asp?TableID={1}&FieldID={0}"
				Case 6
					typeDesc = getAdminSystemLngStr("DtxtCat") & " - " & getAdminSystemLngStr("DtxtStore")
					typeLink = "adminCatOpt.asp?edit=Y&LineIndex={0}&OLKCType=T"
				Case 7
					typeDesc = getAdminSystemLngStr("DtxtCat") & " - " & getAdminSystemLngStr("DtxtCat")
					typeLink = "adminCatOpt.asp?edit=Y&LineIndex={0}&OLKCType=C"
				Case 8
					typeDesc = getAdminSystemLngStr("DtxtCat") & " - " & getAdminSystemLngStr("DtxtList")
					typeLink = "adminCatOpt.asp?edit=Y&LineIndex={0}&OLKCType=L"
				Case 9
					typeDesc = getAdminSystemLngStr("LtxtCMREP")
					typeLink = "adminCartMinRepEdit.asp?goAction=editLine&LineIndex={0}&RowType={1}"
				Case 10
					typeDesc = getAdminSystemLngStr("LtxtObjConfCols") & " - "
					Select Case rs("TypeID")
						Case "A"
							typeDesc = typeDesc & getAdminSystemLngStr("DtxtActions")
						Case "C"
							typeDesc = typeDesc & getAdminSystemLngStr("DtxtClients")
						Case "D"
							typeDesc = typeDesc & getAdminSystemLngStr("DtxtComDocs")
						Case "I"
							typeDesc = typeDesc & getAdminSystemLngStr("DtxtItems")
						Case "R"
							typeDesc = typeDesc & getAdminSystemLngStr("DtxtReceipts")
					End Select
					typeLink = "adminObjConfCols.asp?TypeID={1}&ID={0}"
				Case 11
					typeDesc = getAdminSystemLngStr("DtxtReport") & " ("
					Select Case rs("TypeID")
						Case "C"
							typeDesc = typeDesc & getAdminSystemLngStr("DtxtClient")
						Case "V"
							typeDesc = typeDesc & getAdminSystemLngStr("DtxtAgent")
					End Select
					typeDesc = typeDesc & ")"
					typeLink = "adminRepEdit.asp?rsIndex={0}"
				Case 13
					typeDesc = getAdminSystemLngStr("DtxtFlow")
					
					Select Case rs("TypeID")
						Case "O1"
							typeDescAdd = getAdminSystemLngStr("DtxtAction") & " - " & getAdminSystemLngStr("LtxtConvQuoteOrder")
						Case "O7" 
							typeDescAdd = getAdminSystemLngStr("DtxtAction") & " - " & getAdminSystemLngStr("LtxtConvOrderInv")
						Case "O0" 
							typeDescAdd = getAdminSystemLngStr("DtxtAction") & " - " & getAdminSystemLngStr("LtxtAprovOrder")
						Case "O2" 
							typeDescAdd = getAdminSystemLngStr("DtxtAction") & " - " & getAdminSystemLngStr("LtxtCloseObj")
						Case "O3"
							typeDescAdd = getAdminSystemLngStr("DtxtAction") & " - " & getAdminSystemLngStr("LtxtCancelObj")
						Case "O4" 
							typeDescAdd = getAdminSystemLngStr("DtxtAction") & " - " & getAdminSystemLngStr("LtxtRemObj")
						Case "D1" 
							typeDescAdd = getAdminSystemLngStr("DtxtComDocs") & " - " & getAdminSystemLngStr("LtxtDocCreation")
						Case "D2" 
							typeDescAdd = getAdminSystemLngStr("DtxtComDocs") & " - " & getAdminSystemLngStr("LtxtAddItem")
						Case "D3" 
							typeDescAdd = getAdminSystemLngStr("DtxtComDocs") & " - " & getAdminSystemLngStr("DtxtAdd") & "/" & getAdminSystemLngStr("LtxtDocConf")
						Case "R1" 
							typeDescAdd = getAdminSystemLngStr("DtxtReceipts") & " - " & getAdminSystemLngStr("LtxtCreation")
						Case "R2" 
							typeDescAdd = getAdminSystemLngStr("DtxtReceipts") & " - " & getAdminSystemLngStr("DtxtAdd") & "/" & getAdminSystemLngStr("LtxtRcpConf")
						Case "A1" 
							typeDescAdd = getAdminSystemLngStr("DtxtItems") & " - " & getAdminSystemLngStr("DtxtAdd") & "/" & getAdminSystemLngStr("LtxtItmConf")
						Case "C1" 
							typeDescAdd = getAdminSystemLngStr("DtxtClients") & " - " & getAdminSystemLngStr("DtxtAdd") & "/" & getAdminSystemLngStr("LtxtClientConf")
						Case "C2" 
							typeDescAdd = getAdminSystemLngStr("DtxtClients") & " - " & getAdminSystemLngStr("DtxtAdd") & "/" & getAdminSystemLngStr("LtxtActivityConf")
						Case "C3" 
							typeDescAdd = getAdminSystemLngStr("DtxtClients") & " - " & getAdminSystemLngStr("DtxtAdd") & "/" & getAdminSystemLngStr("LtxtSOConf")
					End Select
					typeDesc = typeDesc & " - " & typeDescAdd
					typeLink = "adminDocFlow.asp?FlowID={0}"
			End Select
			Select Case rs("IsValid")
				Case "U"
					statusDesc = getAdminSystemLngStr("DtxtUnknown")
					statusImg = "question"
				Case "Y"
					statusDesc = getAdminSystemLngStr("DtxtYes")
					statusImg = "check"
				Case "N"
					statusDesc = getAdminSystemLngStr("DtxtError") & ": " & rs("ErrMessage")
					statusImg = "error_db"
			End Select %>
			<tr class="TblRepTbl" style="height: 16px;">
				<td width="10"><a href="<%=Replace(Replace(typeLink, "{0}", rs("ID")), "{1}", rs("TypeID"))%>"><img border="0" src="images/<%=Session("rtl")%>flechaselec.gif" width="15" height="13"></a></td>
				<td><%=typeDesc%></td>
				<td><%=rs("Name")%></td>
				<td id="txtLC_<%=rs("Type")%>_<%=rs("TypeID")%>_<%=Replace(rs("ID"), "-", "_")%>_<%=Replace(rs("ID2"), "-", "_")%>"><%=rs("LastCheck")%></td>
				<td align="center"><img id="img_<%=rs("Type")%>_<%=rs("TypeID")%>_<%=Replace(rs("ID"), "-", "_")%>_<%=Replace(rs("ID2"), "-", "_")%>" src="images/<%=statusImg%>.gif" alt="<%=statusDesc%>">
				<input type="hidden" name="Type" value="<%=rs("Type")%>">
				<input type="hidden" name="TypeID" value="<%=rs("TypeID")%>">
				<input type="hidden" name="FieldID" value="<%=rs("ID")%>">
				<input type="hidden" name="FieldID2" value="-1">
				</td>
			</tr>
			<% Select Case rs("Type") 
				Case 11 
				rd.Filter = "ID = " & rs("ID")
				do while not rd.eof
				Select Case rd("Type")
					Case 12
						typeDesc = getAdminSystemLngStr("DtxtReport") & " ("
						Select Case rd("TypeID")
							Case "C"
								typeDesc = typeDesc & getAdminSystemLngStr("DtxtClient")
							Case "V"
								typeDesc = typeDesc & getAdminSystemLngStr("DtxtAgent")
						End Select
						typeDesc = typeDesc & ") - Values"
						typeLink = "adminRepEdit.asp?rsIndex={0}&editIndex={2}&repCmd=variables&#editVar"
				End Select
				Select Case rd("IsValid")
					Case "U"
						statusDesc = getAdminSystemLngStr("DtxtUnknown")
						statusImg = "question"
					Case "Y"
						statusDesc = getAdminSystemLngStr("DtxtYes")
						statusImg = "check"
					Case "N"
						statusDesc = getAdminSystemLngStr("DtxtError") & ": " & rd("ErrMessage")
						statusImg = "error_db"
				End Select %>
			<tr class="TblRepTbl" style="height: 16px;">
				<td width="10"><a href="<%=Replace(Replace(Replace(typeLink, "{0}", rd("ID")), "{1}", rd("TypeID")), "{2}", rd("ID2"))%>"><img border="0" src="images/<%=Session("rtl")%>flechaselec.gif" width="15" height="13"></a></td>
				<td><%=typeDesc%></td>
				<td><%=rd("Name")%></td>
				<td id="txtLC_<%=rd("Type")%>_<%=rd("TypeID")%>_<%=Replace(rd("ID"), "-", "_")%>_<%=Replace(rd("ID2"), "-", "_")%>"><%=rd("LastCheck")%></td>
				<td align="center"><img id="img_<%=rd("Type")%>_<%=rd("TypeID")%>_<%=Replace(rd("ID"), "-", "_")%>_<%=Replace(rd("ID2"), "-", "_")%>" src="images/<%=statusImg%>.gif" alt="<%=statusDesc%>">
				<input type="hidden" name="Type" value="<%=rd("Type")%>">
				<input type="hidden" name="TypeID" value="<%=rd("TypeID")%>">
				<input type="hidden" name="FieldID" value="<%=rd("ID")%>">
				<input type="hidden" name="FieldID2" value="<%=rd("ID2")%>">
				</td>
			</tr>
			<% 
				rd.movenext
				loop 
			Case 13
				rf.Filter = "TypeID = '" & rs("TypeID") & "' and ID = " & rs("ID")
				do while not rf.eof
				Select Case rf("Type")
					Case 14
						typeDesc = getAdminSystemLngStr("DtxtFlow") & " - " & typeDescAdd & " - " & getAdminSystemLngStr("DtxtMessage")
						typeLink = "adminDocFlow.asp?FlowID={0}"
					Case 15
						typeDesc = getAdminSystemLngStr("DtxtFlow") & " - " & typeDescAdd & " - " & getAdminSystemLngStr("DtxtDetail")
						typeLink = "adminDocFlow.asp?FlowID={0}"
				End Select
				Select Case rf("IsValid")
					Case "U"
						statusDesc = getAdminSystemLngStr("DtxtUnknown")
						statusImg = "question"
					Case "Y"
						statusDesc = getAdminSystemLngStr("DtxtYes")
						statusImg = "check"
					Case "N"
						statusDesc = getAdminSystemLngStr("DtxtError") & ": " & rf("ErrMessage")
						statusImg = "error_db"
				End Select %>
			<tr class="TblRepTbl" style="height: 16px;">
				<td width="10"><a href="<%=Replace(Replace(Replace(typeLink, "{0}", rf("ID")), "{1}", rf("TypeID")), "{2}", rf("ID2"))%>"><img border="0" src="images/<%=Session("rtl")%>flechaselec.gif" width="15" height="13"></a></td>
				<td><%=typeDesc%></td>
				<td><%=rf("Name")%></td>
				<td id="txtLC_<%=rf("Type")%>_<%=rf("TypeID")%>_<%=Replace(rf("ID"), "-", "_")%>_<%=Replace(rf("ID2"), "-", "_")%>"><%=rf("LastCheck")%></td>
				<td align="center"><img id="img_<%=rf("Type")%>_<%=rf("TypeID")%>_<%=Replace(rf("ID"), "-", "_")%>_<%=Replace(rf("ID2"), "-", "_")%>" src="images/<%=statusImg%>.gif" alt="<%=statusDesc%>">
				<input type="hidden" name="Type" value="<%=rf("Type")%>">
				<input type="hidden" name="TypeID" value="<%=rf("TypeID")%>">
				<input type="hidden" name="FieldID" value="<%=rf("ID")%>">
				<input type="hidden" name="FieldID2" value="<%=rf("ID2")%>">
				</td>
			</tr>
			<% 
				rf.movenext
				loop
			End Select
			rs.movenext
			loop %>
		</table>
		</div>
		</div>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr>
				<td width="77">
				<input type="submit" value="<%=getadminSystemLngStr("LtxtTest")%>" name="btnTest" class="OlkBtn" onclick="startTest();"></td>
				<td><hr color="#0D85C6" size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
	</form>
</table>
<script type="text/javascript">
var next = -1;
var errCount = 0;

function startTest()
{
	document.frmTest.btnTest.disabled = true;
	if (document.frmTest.Type)
	{
		if (document.frmTest.Type.length)
		{
			for (var i = 0;i<document.frmTest.Type.length;i++)
			{
				setTestImg(document.frmTest.Type[i].value, document.frmTest.TypeID[i].value, document.frmTest.FieldID[i].value, document.frmTest.FieldID2[i].value, 3);
			}
			next = 1;
			executeTest(document.frmTest.Type[0].value, document.frmTest.TypeID[0].value, document.frmTest.FieldID[0].value, document.frmTest.FieldID2[0].value);
			
		}
		else
		{
			setTestImg(document.frmTest.Type.value, document.frmTest.TypeID.value, document.frmTest.FieldID.value, document.frmTest.FieldID2.value, 3);
			executeTest(document.frmTest.Type.value, document.frmTest.TypeID.value, document.frmTest.FieldID.value, document.frmTest.FieldID2.value);
			finishTest();
		}
	}
}

function finishTest()
{
	if (errCount == 0)
	{
		alert('<%=getadminSystemLngStr("LtxtTestOK")%>');
	}
	else
	{
		alert('The test has finished with {0} errors'.replace('{0}', errCount));
	}
	document.frmTest.btnTest.disabled = false;
}

function executeTest(Type, TypeID, ID, ID2)
{
	$.post("verfyTest.asp?d=" + (new Date()).toString(), { Type: Type, TypeID: TypeID, ID: ID, ID2: ID2 },
   function(data){
   		switch (parseInt(Type))
   		{
   			case 9:
   				var arrAn = data.split('{Q}');
   				for (var i = 0;i<arrAn.length;i++)
   				{
	   				var arrData = arrAn[i].split('{S}');
	   				switch (arrData[0])
	   				{
	   					case 'ok':
	   						setTestImg(Type, TypeID, ID, ID2, 1);
	   						break;
	   					case 'err':
	   						setTestImg(Type, TypeID, ID, ID2, 2);
							document.getElementById(('img_' + Type + '_' + TypeID + '_' + ID + '_' + ID2).replace('-', '_')).alt = 'Error: ' + arrData[2];
							i++;
							errCount++;
	   						break;
	   				}
   				}
   				break;
   			default:
   				var arrData = data.split('{S}');
   				switch (arrData[0])
   				{
   					case 'ok':
   						setTestImg(Type, TypeID, ID, ID2, 1);
   						break;
   					case 'err':
   						setTestImg(Type, TypeID, ID, ID2, 2);
						document.getElementById('img_' + Type + '_' + TypeID + '_' + ID.replace('-', '_') + '_' + ID2.replace('-', '_')).alt = '<%=getAdminSystemLngStr("DtxtError")%>: ' + arrData[2];
						errCount++;
   						break;
   				}
   				break;
   		}
   		if (next != -1)
   		{
   			if (next < document.frmTest.Type.length)
   			{
   				var i = next++;
   				executeTest(document.frmTest.Type[i].value, document.frmTest.TypeID[i].value, document.frmTest.FieldID[i].value, document.frmTest.FieldID2[i].value);
   			}
   			else
   			{
   				finishTest();
   			}
   		}
   });
}

function setTestImg(Type, TypeID, ID, ID2, imgID)
{
	var img = '';
	switch (imgID)
	{
		case 1:
			img = 'check';
			break;
		case 2:
			img = 'error_db';
			break;
		case 3:
			img = 'ajax-loader';
			break;
	}
	document.getElementById('img_' + Type + '_' + TypeID + '_' + ID.replace('-', '_') + '_' + ID2.replace('-', '_')).src = 'images/' + img + '.gif';
	document.getElementById('img_' + Type + '_' + TypeID + '_' + ID.replace('-', '_') + '_' + ID2.replace('-', '_')).alt = '';
	
	if (imgID != 3)
	{
		$('#divTest').scrollTop($('#img_' + Type + '_' + TypeID + '_' + ID.replace('-', '_') + '_' + ID2.replace('-', '_')).offset().top - $('#divTest').offset().top + $('#divTest').scrollTop());
	}
	
}
</script>
<table border="0" cellpadding="0" width="98%">
<form name="frmRestore" method="post" action="restore.asp">
	<tr>
		<td height="15"></td>
	</tr>
	<tr>
		<td bgcolor="#E1F3FD"><b><font color="#31659C" size="1" face="Verdana">&nbsp;<%=getadminSystemLngStr("LttlRestore")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif">
		<font face="Verdana" color="#4783C5" size="1"><%=getadminSystemLngStr("LttlRestoreNote")%></font></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr>
				<td width="77">
				<input type="submit" value="<%=getadminSystemLngStr("LttlRestore")%>" name="btnRestore" class="OlkBtn" onclick="this.disabled=true;document.frmRestore.submit();"></td>
				<td><hr color="#0D85C6" size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
</form>
</table>
<% End If %><!--#include file="bottom.asp" -->
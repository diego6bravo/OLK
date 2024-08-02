<!--#include file="top.asp" -->
<!--#include file="lang/adminDefObjEdit.asp" -->
<% conn.execute("use [" & Session("OLKDB") & "]") 
ObjID = CInt(Request("ObjID"))
If ObjID <> 19 Then
	sql = "Select SelDes from OLKCommon"
	set rs = conn.execute(sql)
	SelDes = rs(0)
	rs.close
Else
	SelDes = 0
End If %>
<script language="javascript">
function chkNum(fld, oldVal)
{
	if (!IsNumeric(fld.value))
	{
		alert("<%=getadminDefObjEditLngStr("DtxtValNumVal")%>");
		fld.value = oldVal.value;
	}
	else if (fld.value < 0)
	{
		alert("<%=getadminDefObjEditLngStr("DtxtValNumMinVal")%>".replace("{0}", "0"));
		fld.value = 0;
	}
	else if (fld.value > 32767)
	{
		alert("<%=getadminDefObjEditLngStr("DtxtValNumMaxVal")%>".replace("{0}", "32767"));
		fld.value = 32767;
	}
	oldVal.value = fld.value;
}
</script>


<table border="0" cellpadding="0" width="100%" id="table3">
	<tr>
		<td>&nbsp;</td>
	</tr>
	<form method="POST" action="adminSubmit.asp" name="frmAddEditObj">
		<tr>
			<td bgcolor="#E1F3FD"><b><font face="Verdana" size="2">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><% If ObjID = "" Then %><%=getadminDefObjEditLngStr("LtxtAddObj")%><% Else %><%=getadminDefObjEditLngStr("LtxtEditObj")%><% End If %></font></b></td>
		</tr>
		<tr>
			<td bgcolor="#F5FBFE">
			<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1">
			<font color="#4783C5"><%=getadminDefObjEditLngStr("LtxtObjNote")%></font></font></p>
			</td>
		</tr>
		<%
	set rs = nothing
	If ObjID <> "" Then
		sql = "select Case T0.ObjType When 'S' Then T1.ObjName collate database_default Else T0.ObjName End ObjName, " & _
		"T0.ObjContent, T0.Status " & _
		"from OLKObjects T0 " & _
		"left outer join OLKCommon..OLKObjectsDesc T1 on T1.ObjID = T0.ObjID and T0.ObjType = 'S' and T1.LanID = " & Session("LanID") & " " & _
		"where T0.ObjType = '" & Request("ObjType") & "' and T0.ObjID = " & ObjID
		set rs = conn.execute(sql)
		ObjName = rs("ObjName")
		Status = rs("Status")
		ObjContent = rs("ObjContent")
		set rd = Server.CreateObject("ADODB.RecordSet")
		sql = "select * from OLKObjectsVars where ObjType = '" & Request("ObjType") & "' and ObjId = " & ObjID
		rd.open sql, conn, 3, 1
	Else
		Status = "N"
		NewContent = "<font face=""verdana"" size=""1""><div></div></font>"
		ObjContent = ""
	End If %>
		<tr>
			<td>
			<table border="0" cellpadding="0" width="100%" id="table11">
				<tr bgcolor="#E2F3FC">
					<td align="center"><b>
					<font face="Verdana" size="1" color="#31659C"><%=getadminDefObjEditLngStr("DtxtName")%></font></b></td>
					<td align="center" style="width: 15%"><b>
					<font face="Verdana" size="1" color="#31659C"><%=getadminDefObjEditLngStr("DtxtActive")%></font></b></td>
				</tr>
				<tr>
					<td bgcolor="#F5FBFE">
					<font face="Verdana" size="1" color="#4783C5"><%=rs("ObjName")%></font>&nbsp;</td>
					<td bgcolor="#F5FBFE" style="width: 15%">
					<p align="center">
					<input type="checkbox" name="Status" <% if Status = "Y" then %>checked<% end if %> value="Y" class="noborder"></p>
					</td>
				</tr>
				<% If not rd is nothing then
				If not rd.eof then %>
				<tr bgcolor="#E2F3FC">
					<td colspan="2">
					<p align="center"><b>
					<font face="Verdana" size="1" color="#31659C">
					<%=getadminDefObjEditLngStr("DtxtVariables")%></font></b></p>
					</td>
				</tr>
				<tr bgcolor="#E2F3FC">
					<td colspan="2" bgcolor="white">
					<table border="0" id="table12" cellpadding="0" style="width: 50%">
						<tr>
							<td bgcolor="#E2F3FC">
							<p align="center"><b>
							<font face="Verdana" size="1" color="#31659C">
							<%=getadminDefObjEditLngStr("DtxtName")%></font></b></p>
							</td>
							<td bgcolor="#E2F3FC">
							<p align="center"><b>
							<font face="Verdana" size="1" color="#31659C">
							<%=getadminDefObjEditLngStr("DtxtType")%></font></b></p>
							</td>
							<td bgcolor="#E2F3FC">
							<p align="center"><b>
							<font face="Verdana" size="1" color="#31659C">
							<%=getadminDefObjEditLngStr("DtxtValue")%></font></b></p>
							</td>
						</tr>
						<% do while not rd.eof %>
						<tr>
							<td bgcolor="#F5FBFE">
							<font face="Verdana" size="1" color="#31659C"><%=rd("VarName")%>
							</font></td>
							<td bgcolor="#F5FBFE">
							<font face="Verdana" size="1" color="#31659C"><% Select Case rd("VarType")
							Case "N" %>
							<%=getadminDefObjEditLngStr("DtxtNumeric")%>
							<% Case "A" %>
							<%=getadminDefObjEditLngStr("DtxtAlphaNumeric")%>
							<% End Select %></font></td>
							<td bgcolor="#F5FBFE">
							<input type="text" style="width: 100%;" name="VarVal<%=rd("VarId")%>" size="20" style="<% If rd("VarType") = "N" Then %>text-align:right;<% End If %>" value="<%=rd("VarValue")%>" onfocus="this.select()" onchange="<% If rd("VarType") = "N" Then %>chkNum(this, document.frmAddEditObj.oldVarVal<%=rd("VarId")%>);<% End If %>newPreview();">
							<% If rd("VarType") = "N" Then %><input type="hidden" name="oldVarVal<%=rd("VarId")%>" value="<%=rd("VarValue")%>"><% End If %></td>
						</tr>
						<% rd.movenext
					loop %>
					</table>
					</td>
				</tr>
				<% End If
			End If %>
				<tr bgcolor="#E2F3FC">
					<td colspan="2">
					<p align="center"><b>
					<font face="Verdana" size="1" color="#31659C"><%=getadminDefObjEditLngStr("LtxtContent")%></font></b></p>
					</td>
				</tr>
				<tr>
					<td colspan="2">
					<table border="0" id="table13" cellpadding="0">
						<tr>
							<td id="btnPrevContent" width="100" bgcolor="#D9F5FF" align="center" style="border: 1px solid #31659C" onclick="showPrevContent();" onmouseover="javascript:if(document.getElementById('PrevContent').style.display=='none')this.bgColor='#D9F5FF';" onmouseout="javascript:if(document.getElementById('PrevContent').style.display=='none')this.bgColor='#BFEEFE';">
							<font color="#31659C" face="Verdana" size="1"><b>
							<%=getadminDefObjEditLngStr("DtxtPreview")%></b></font></td>
							<td id="btnEditContent" width="60" bgcolor="#BFEEFE" align="center" style="border: 1px solid #31659C; cursor: hand" onclick="showEditContent();" onmouseover="javascript:if(document.frmAddEditObj.ObjContent.style.display=='none')this.bgColor='#D9F5FF';" onmouseout="javascript:if(document.frmAddEditObj.ObjContent.style.display=='none')this.bgColor='#BFEEFE';">
							<font color="#31659C" face="Verdana" size="1"><b>
							<%=getadminDefObjEditLngStr("DtxtEdit")%></b></font></td>
						</tr>
					</table>
					</td>
				</tr>
				<tr bgcolor="#E2F3FC">
					<td colspan="2">
					<p align="center">
					<iframe name="PrevContent" id="PrevContent" src="clear.asp" target="_blank" style="width: 100%; height: 580px">
					Your browser does not support inline frames or is currently 
					configured not to display inline frames.
					</iframe>
					<textarea id="ObjContent" name="ObjContent" style="display: none; width: 100%; height: 580px;" cols="1" dir="ltr"><%=Replace(ObjContent, "&", "&amp;")%></textarea></p>
					</td>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td>
			<table border="0" cellpadding="0" width="100%" id="table5">
				<tr>
					<td width="77">
					<input type="submit" value="<%=getadminDefObjEditLngStr("DtxtApply")%>" name="btnApply" class="OlkBtn"></td>
					<td width="77">
					<input type="submit" value="<%=getadminDefObjEditLngStr("DtxtSave")%>" name="btnSave" class="OlkBtn"></td>
					<td><hr color="#0D85C6" size="1"></td>
					<% If Request("ObjType") = "S" Then %>
					<td width="77">
					<input type="submit" value="<%=getadminDefObjEditLngStr("DtxtRestore")%>" name="btnRestore" class="OlkBtn" onclick="javascript:return confirm('<%=getadminDefObjEditLngStr("LtxtValRestoreFld")%>');"></td><% End If %>
					<td width="77">
					<input type="button" value="<%=getadminDefObjEditLngStr("DtxtCancel")%>" name="B1" class="OlkBtn" onclick="javascript:if(confirm('<%=getadminDefObjEditLngStr("DtxtConfCancel")%>'))window.location.href='adminDefObjs.asp'"></td>
				</tr>
			</table>
			</td>
		</tr>
		<input type="hidden" name="submitCmd" value="adminObjs">
		<input type="hidden" name="uCmd" value="edit">
		<input type="hidden" name="ObjType" value="<%=Request("ObjType")%>">
		<input type="hidden" name="ObjId" value="<%=ObjID%>">
	</form>
	<tr>
		<td>&nbsp;</td>
	</tr>
</table>
<script language="javascript">
function newPreview()
{
	strPrev = document.frmAddEditObj.ObjContent.value;
	//strPrev = strPrev.replace('src="', 'src=\"')
	strPrev = strPrev.replace('{dbName}', '<%=Session("olkdb")%>')
	strPrev = strPrev.replace('{SelDes}', '<%=SelDes%>');
	strPrev = strPrev.replace('{rtl}', '<% If Session("rtl") <> "" Then Response.Write "rtl" Else Response.Write "" %>');
	strPrev = strPrev.replace('{rtl2}', '<%=Session("rtl")%>');
<%	If ObjID = 8 or ObjID = 7 Then %>
	strPrev = strPrev.replace('<!--startPicLink-->', '');
	strPrev = strPrev.replace('<!--endPicLink-->', '');
	strPrev = strPrev.replace('<!--startPicMoreLink-->', '');
	strPrev = strPrev.replace('<!--endPicMoreLink-->', '');
	strPrev = strPrev.replace('{Picture}', 'n_a.gif');
	strPrev = strPrev.replace('{ImgMaxSize}', 100);
	strPrev = strPrev.replace('{ItemType}', 'item');
<%
	End If
	If Not rd is nothing Then
		If rd.recordcount > 0 Then
			rd.movefirst
			do while not rd.eof %>
			strPrev = strPrev.replace('{<%=rd("VarName")%>}', document.frmAddEditObj.VarVal<%=rd("VarId")%>.value);
	<% 		rd.movenext
			loop
		End If
	End If %>
	
	doPreview(strPrev);
}

function showPrevContent()
{
	if (document.getElementById('PrevContent').style.display == 'none')
	{
		newPreview();
		
		document.getElementById('PrevContent').style.display = '';
		document.getElementById('ObjContent').style.display = 'none';

		btnPrevContent.bgColor = '#D9F5FF';
		btnPrevContent.style.cursor = '';
		btnEditContent.bgColor = '#BFEEFE';
		btnEditContent.style.cursor = 'hand';
	}
}
function showEditContent()
{
	if (document.getElementById('ObjContent').style.display == 'none')
	{
		document.getElementById('PrevContent').style.display = 'none';
		document.getElementById('ObjContent').style.display = '';
		
		btnPrevContent.bgColor = '#BFEEFE';
		btnPrevContent.style.cursor = 'hand';
		btnEditContent.bgColor = '#D9F5FF';
		btnEditContent.style.cursor = '';
	}
}
function doPreview(Content)
{
	<% If ObjID = 19 Then %>Content = '<table style="width: 100%; background-color: white;">' + Content + '</table>';<% End If %>
	PrevContent.document.body.innerHTML = Content;
}

function doLoadPreview()
{
	PrevContent.changeStyle('design/<%=SelDes%>/style/stylenuevo.css');
	<% strPreview = Replace(ObjContent, "'", "\'")
	strPreview = Replace(strPreview, VbNewLine, "\n")
	'strPreview = Replace(strPreview, "src=""", "src=""")
	strPreview = Replace(strPreview, "{dbName}", Session("olkdb"))
	strPreview = Replace(strPreview, "{SelDes}", SelDes)
	strPreview = Replace(strPreview, "{rtl}", Session("rtl"))
	If Session("rtl") <> "" Then
		strPreview = Replace(strPreview, "{rtl2}", "rtl")
	Else
		strPreview = Replace(strPreview, "{rtl2}", "")
	End If
	If ObjID = 8 or ObjID = 7 Then
		strPreview = Replace(strPreview, "<!--startPicLink-->", "")
		strPreview = Replace(strPreview, "<!--endPicLink-->", "")
		strPreview = Replace(strPreview, "<!--startPicMoreLink-->", "")
		strPreview = Replace(strPreview, "<!--endPicMoreLink-->", "")
		strPreview = Replace(strPreview, "{Picture}", "n_a.gif")
		strPreview = Replace(strPreview, "{ImgMaxSize}", 100)
		strPreview = Replace(strPreview, "{ItemType}", "item")
	End If
	If not rd is nothing Then
		If rd.recordcount > 0 Then
			rd.movefirst
			do while not rd.eof
				strPreview = Replace(strPreview, "{" & rd("VarName") & "}", rd("VarValue"))
			rd.movenext
			loop
		End If
	End If
	
	If ObjID = 19 Then strPreview = "<table style=""width: 100%; background-color: white;"">" & strPreview & "</table>"
	%>
	doPreview('<%=strPreview%>');
}
</script>
<!--#include file="bottom.asp" -->
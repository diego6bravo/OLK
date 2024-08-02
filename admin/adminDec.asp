<!--#include file="top.asp" -->
<!--#include file="lang/adminDec.asp" -->


<% conn.execute("use [" & Session("OLKDB") & "]")
If Request.Form.Count > 0 Then
	sql = "update OLKCommon set SelDes = N'" & Request("SelDes") & "'"
	conn.execute(sql)
End If
sql = "select SelDes from OLKCommon"
set rs = conn.execute(sql)
selDes = rs(0)
rs.close

sql = "select DisID, AlterName from OLKCommon..OLKCustDes"
rs.open sql, conn, 3, 1 %>
<br>
<table border="0" cellpadding="0" width="100%" id="table3">
<form method="POST" action="adminDec.asp">
	<tr>
		<td bgcolor="#E7F3FF">&nbsp;<b><font face="Verdana" size="1" color="#31659C"><%=getadminDecLngStr("LttlSelDes")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif">
		<font face="Verdana" size="1" color="#4783C5"><%=getadminDecLngStr("LttlSelDesNote")%></font></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<table border="0" cellpadding="0" width="100%" id="table4">
				<tr>
					<td width="33%" align="center">
				<a target="_blank" href="desPrev/prev1.jpg">
				<img border="1" src="prev.aspx?filename=desPrev/prev1.jpg&amp;MaxSize=150" style="border: 1px solid #000000"></a></td>
					<td width="33%" align="center">
				<a href="desPrev/prev2.jpg" target="_blank">
				<img border="1" src="prev.aspx?filename=desPrev/prev2.jpg&amp;MaxSize=150" style="border: 1px solid #000000"></a></td>
					<td width="33%" align="center">
				<a target="_blank" href="desPrev/prev3.jpg">
				<img border="1" src="prev.aspx?filename=desPrev/prev3.jpg&amp;MaxSize=150" style="border: 1px solid #000000"></a></td>
				</tr>
				<tr>
					<td width="33%" align="center">
					<font color="#4783C5" face="Verdana" size="1">
					<%=getadminDecLngStr("LtxtHomeLogo")%> 248x122<br>
					<%=getadminDecLngStr("LtxtMailLogo")%> 247x164</font></td>
					<td align="center" width="33%">
					<font color="#4783C5" face="Verdana" size="1">
					<%=getadminDecLngStr("LtxtHomeLogo")%> 248x88<br>
					<%=getadminDecLngStr("LtxtMailLogo")%> 247x164</font></td>
					<td align="center" width="33%">
					<font color="#4783C5" face="Verdana" size="1">
					<%=getadminDecLngStr("LtxtHomeLogo")%> 248x133<br>
					<%=getadminDecLngStr("LtxtMailLogo")%> 247x164</font></td>
				</tr>
				<tr>
					<td width="33%" align="center">
					<b>
					<font color="#4783C5" face="Verdana" size="1">
					<input type="radio" value="1" <% If selDes = "1" Then %>checked<% End If %> name="SelDes" id="rdDef1" style="border-style:solid; border-width:0; background:background-image"><label for="rdDef1"><br>
					Classic</label></font></b></td>
					<td align="center" width="33%">
					<b>
					<font color="#4783C5" face="Verdana" size="1">
					<input type="radio" name="SelDes" <% If selDes = "2" Then %>checked<% End If %> value="2" id="rdDef2" style="border-style:solid; border-width:0; background:background-image"><label for="rdDef2"><br>
					Summer</label></font></b></td>
					<td align="center" width="33%">
					<b>
					<font color="#4783C5" face="Verdana" size="1">
					<input type="radio" name="SelDes" <% If selDes = "3" Then %>checked<% End If %> value="3" id="rdDef3" style="border-style:solid; border-width:0; background:background-image"><label for="rdDef3"><br>
					Cool Blue</label></font></b></td>
				</tr>
				<tr>
					<td width="33%">
					&nbsp;</td>
					<td align="center" width="33%">
					&nbsp;</td>
					<td align="center" width="33%">&nbsp;</td>
				</tr>
				<tr>
					<td width="33%">
					<p align="center">
				<a target="_blank" href="desPrev/prev4.jpg">
				<img border="1" src="prev.aspx?filename=desPrev/prev4.jpg&amp;MaxSize=150" style="border: 1px solid #000000"></a></td>
					<td align="center" width="33%">
					<a target="_blank" href="desPrev/prev5.jpg">
				<img border="1" src="prev.aspx?filename=desPrev/prev5.jpg&amp;MaxSize=150" style="border: 1px solid #000000"></a></td>
					<td align="center" width="33%">
					<a target="_blank" href="desPrev/prev6.jpg">
				<img border="1" src="prev.aspx?filename=desPrev/prev6.jpg&amp;MaxSize=150" style="border: 1px solid #000000"></a></td>
				</tr>
				<tr>
					<td width="33%">
					<p align="center">
					<font color="#4783C5" face="Verdana" size="1">
					<%=getadminDecLngStr("LtxtHomeLogo")%> 226x136<br>
					<%=getadminDecLngStr("LtxtMailLogo")%> 247x164</font></td>
					<td align="center" width="33%">
					<p align="center">
					<font color="#4783C5" face="Verdana" size="1">
					<%=getadminDecLngStr("LtxtHomeLogo")%> 251x105<br>
					<%=getadminDecLngStr("LtxtMailLogo")%> 247x164</font></td>
					<td align="center" width="33%">
					<p align="center">
					<font color="#4783C5" face="Verdana" size="1">
					<%=getadminDecLngStr("LtxtHomeLogo")%> 166x86<br>
					<%=getadminDecLngStr("LtxtMailLogo")%> 247x164</font></td>
				</tr>
				<tr>
					<td width="33%">
					<p align="center">
					<b>
					<font color="#4783C5" face="Verdana" size="1">
					<input type="radio" value="4" <% If selDes = "4" Then %>checked<% End If %> name="SelDes" id="rdDef4" style="border-style:solid; border-width:0; background:background-image"><label for="rdDef4"><br>
					Hi-Tech</label></font></b></td>
					<td align="center" width="33%">
					<p align="center">
					<b>
					<font color="#4783C5" face="Verdana" size="1">
					<input type="radio" value="5" <% If selDes = "5" Then %>checked<% End If %> name="SelDes" id="rdDef5" style="border-style:solid; border-width:0; background:background-image"><label for="rdDef5"><br>
					Sunshine</label></font></b></td>
					<td align="center" width="33%">
					<p align="center">
					<b>
					<font color="#4783C5" face="Verdana" size="1">
					<input type="radio" value="6" <% If selDes = "6" Then %>checked<% End If %> name="SelDes" id="rdDef6" style="border-style:solid; border-width:0; background:background-image"><label for="rdDef6"><br>
					Caribbean Red</label></font></b></td>
				</tr>
				<tr>
					<td width="33%">&nbsp;</td>
					<td align="center" width="33%">
					&nbsp;</td>
					<td align="center" width="33%">&nbsp;</td>
				</tr>
			</table>
		</td>
	</tr>
	<!--#include file="design/customid.inc" -->
	<% If EnCostDis Then %>
	<tr>
		<td bgcolor="#E7F3FF">&nbsp;<b><font face="Verdana" size="1" color="#31659C"><%=getadminDecLngStr("LttlPerDes")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<table border="0" cellpadding="0" width="100%" id="table4">
				<tr>
				<% 
				curDis = 0 
				varx = 0
				Dim CustDisId(2)
				Dim CustDisNam(2)
				CustDisNam(0) = ""
				CustDisNam(1) = ""
				CustDisNam(2) = ""
				For i = 0 to UBound(CustDis)
					CustDisId(varx) = Split(CustDis(i),",")(0)
					CustDisNam(varx) = Split(CustDis(i),",")(1) %>
					<td width="33%">
					<p align="center">
					<a target="_blank" href="design/custom<%=CustDisId(varx)%>/prev.jpg">
					<img border="1" src="prev.aspx?filename=design/custom<%=CustDisId(varx)%>/prev.jpg&amp;MaxSize=150" style="border: 1px solid #000000"></a></td>
				<% 
				varx = varx + 1
				If varx = 3 or varx < 3 and i = UBound(CustDis) Then
					varx = 0
					Response.Write "</tr><tr>"
					For n = 0 to 2
						If CustDisNam(n) <> "" Then
							If "custom" & CustDisId(n) = SelDes Then chk = "checked" Else chk = ""
							rs.Filter = "DisID = " & CustDisId(n)
							If Not rs.Eof Then CustDisNam(n) = rs("AlterName")
							Response.Write "<td width=""33%""> " & _
											"<p align=""center""><b><font color=""#4783C5"" face=""Verdana"" size=""1""> " & _
											"<input " & chk & " type=""radio"" value=""custom" & CustDisId(n) & """ name=""SelDes"" id=""rdCustDef" & CustDisId(n) & """ style=""border-style:solid; border-width:0; background:background-image""><label for=""rdCustDef" & CustDisId(n) & """><br> " & _
											"<span id=""disNam" & CustDisId(n) & """>" & CustDisNam(n) & "</span></label></font></b><br>" & _
											"<input type=""button"" name=""btnEditDisName"" value=""" & getadminDecLngStr("DtxtEdit") & """  style=""color: #68A6C0; font-family: Tahoma; border: 1px solid #68A6C0; background-color: #E5F1FF; font-size:10px; width:75; height:23; font-weight:bold"" onclick=""doEditDis(" & CustDisId(n) & ", '" & Replace(CustDisNam(n), "'", "\'") & "')""><br>&nbsp;</td> " 
							CustDisId(n) = ""
							CustDisNam(n) = ""
						End If
					Next
				End If
				curDis = curDis + 1
				If curDis = 3 Then
					curDis = 0
					Response.Write "</tr><tr>"
				End If
				Next %>
				</tr>
				<tr>
					<td width="33%">&nbsp;</td>
					<td align="center" width="33%">
					&nbsp;</td>
					<td align="center" width="33%">&nbsp;</td>
				</tr>
		</table>
		</td>
	</tr>
	<% End If %>
	<input type="hidden" name="cmd" value="design">
	<tr>
		<td>
		<table cellpadding="0" cellspacing="0" border="0" width="100%">
			<tr>
				<td width="77"><input type="submit" value="<%=getadminDecLngStr("DtxtSave")%>" name="btnSave" style="color: #68A6C0; font-family: Tahoma; border: 1px solid #68A6C0; background-color: #E5F1FF; font-size:10px; width:77; height:23; font-weight:bold"></td>
				<td><hr size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
	</form>
</table>
<script language="javascript">
function doEditDis(ID, Name)
{
	OpenWin = this.open('adminEditDec.asp?Pop=Y&ID=' + ID + '&Name=' + Name, "OpenWin", "toolbar=no,menubar=no,location=no,scrollbars=no,resizable=no, width=300,height=130");
}
function setDisName(ID, AlterName)
{
	document.getElementById('disNam' + ID).innerHTML = AlterName;
}
</script>
<!--#include file="bottom.asp" -->
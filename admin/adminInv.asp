<!--#include file="top.asp" -->
<!--#include file="lang/adminInv.asp" -->
<!--#include file="adminTradSubmit.asp"-->
<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<script type="text/javascript">
<!--
function chkThis(fld, min, dec, oldVal)
{
	if (!IsNumeric(fld.value))
	{
		alert('<%=getadminInvLngStr("DtxtValNumVal")%>');
		fld.value = oldVal;
		fld.focus();
	}
	else if (parseFloat(fld.value) < parseFloat(min))
	{
		fld.value = FormatNumber(min, dec);
		alert("<%=getadminInvLngStr("DtxtValNumMinVal")%>".replace("{0}", min));
		fld.focus();
	}
	else if (parseFloat(fld.value) > 32727)
	{
		fld.value = 32767;
		alert("<%=getadminInvLngStr("DtxtValNumMaxVal")%>".replace("{0}", "32767"));
		fld.focus();
	}
	fld.value = formatNumber(fld.value, dec).replace(',', '');
}
//-->
</script>
<script language="javascript">
function valFrm()
{
	if (document.form1.valGenFilter.value == 'Y' && document.form1.GenFilter.value != '')
	{
		alert("<%=getadminInvLngStr("LtxtValCatFltQryVal")%>");
		document.form1.GenFilter.focus();
		return false;
	}
	else if ((document.form1.GenFAppC.checked || document.form1.GenFAppV.checked) && document.form1.GenFilter.value == '')
	{
		alert("<%=getadminInvLngStr("LtxtValFldQry")%>");
		document.form1.GenFilter.focus();
		return false;
	}
	else if (document.form1.WhsCode.selectedIndex == 0)
	{
		alert("<%=getadminInvLngStr("LtxtValWhs")%>");
		document.form1.WhsCode.focus();
		return false;
	}
	else if (document.form1.EnableItemRec.checked && document.form1.ItemRecQry.value == '')
	{
		alert('<%=getadminInvLngStr("LtxtValItemRecQry")%>');
		document.form1.ItemRecQry.focus();
		return false;
	}
	else if (document.form1.ItemRecQry.value != '' && document.form1.valItemRecQry.value == 'Y')
	{
		alert('<%=getadminInvLngStr("LtxtValRecItemQry")%>');
		document.form1.ItemRecQry.focus();
		return false;
	}
	else if (document.form1.EnableCodeBarsQry.checked && document.form1.CodeBarsQry.value == '')
	{
		alert('<%=getadminInvLngStr("LtxtValCodeBarQry")%>');
		document.form1.CodeBarsQry.focus();
		return false;
	}
	else if (document.form1.CodeBarsQry.value != '' && document.form1.valCodeBarsQry.value == 'Y')
	{
		alert('<%=getadminInvLngStr("LtxtValBarCodeQry")%>');
		document.form1.CodeBarsQry.focus();
		return false;
	}
	return true;
}
</script>
<script language="javascript" src="js_up_down.js"></script>
<style type="text/css">
.style1 {
	background-color: #E1F3FD;
	font-family: Verdana;
	font-size: xx-small;
	color: #4783C5;
	text-align: center;
}
.style2 {
	font-family: Verdana;
	font-size: xx-small;
	color: #4783C5;
	text-align: center;
}
.style3 {
	font-family: Verdana;
	font-size: xx-small;
	color: #4783C5;
}
.style5 {
				color: #FF0000;
}
</style>
</head>
<% If Session("style") = "nc" Then %>
<br>
<% End If %>
<form name="form1" action="adminsubmit.asp" method="POST" onsubmit="javascript:return valFrm()">
<table border="0" cellpadding="0" width="100%" id="table3">
	<tr>
		<td bgcolor="#E7F3FF">&nbsp;<b><font size="1" face="Verdana" color="#31659C"><%=getadminInvLngStr("LttlInvSearch")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1"> 
		<font color="#4783C5"><%=getadminInvLngStr("LttlInvSearchNote")%></font></font></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<table border="0" cellpadding="0" width="100%" id="table7">
			<tr>
				<td width="300"><img src="images/ganchito.gif"><font face="Verdana" size="1"><span style="font-size: 12px" 12px?? font-size:> </span>
				</font><font face="Verdana" size="1" color="#4783C5">
				<%=getadminInvLngStr("LtxtWhsCode")%></font></td>
				<td>
				<select size="1" name="WhsCode" class="input">
				<option value=""><%=getadminInvLngStr("LoptSelWhs")%></option>
				<%
				GetQuery rd, 1, null, null
				do while not rd.eof %>
				<option <% If CStr(rd("WhsCode")) = CStr(myApp.WhsCode) Then %>selected<%end if%> value="<%=rd("WhsCode")%>"><%=myHTMLEncode(rd("WhsName"))%></option>
				<% rd.movenext
				loop %>
				</select></td>
			</tr>
			<tr>
				<td width="300"><img src="images/ganchito.gif"><font face="Verdana" size="1" color="#4783C5"> <%=getadminInvLngStr("LtxtInvBDGBy")%></font></td>
				<td>
				<font face="Verdana" size="1" color="#4783C5">
				<input type="radio" name="InvBDGBy" value="S" id="fp5" <% If myApp.InvBDGBy = "S" Then Response.write "checked"%> checked class="noborder"><label for="fp5"><%=getadminInvLngStr("DtxtSAP")%></label><input type="radio" name="InvBDGBy" value="E" id="fp4" <% If myApp.InvBDGBy = "E" Then Response.write "checked"%> class="noborder"><label for="fp4"><%=getadminInvLngStr("LtxtCustom")%></label></font></td>
			</tr>
			<tr>
				<td width="300">
				<img src="images/ganchito.gif"><font face="Verdana" size="1" color="#4783C5">&nbsp;</font><input class="noborder" type="checkbox" name="EnableMinInv" id="EnableMinInv" value="Y" <% If myApp.EnableMinInv Then %>checked<% End If %>><font face="Verdana" size="1" color="#4783C5"><label for="EnableMinInv"><%=getadminInvLngStr("LtxtMinInv")%></label></font></td>
				<td>
				<span style="font-size: 12px">
    			<input type="text" name="MinInv" size="10" style="text-align:right" class="input" value="<%=FormatNumber(myApp.MinInv, myApp.QtyDec)%>" onfocus="this.select()" onchange="chkThis(this, 0, <%=myApp.QtyDec%>, document.form1.oldMinInv.value);document.form1.oldMinInv.value=this.value;" onkeydown="return chkMax(event, this, 6);">
    			<input type="hidden" name="oldMinInv" value="<%=FormatNumber(myApp.MinInv, myApp.QtyDec)%>">
				</span><font face="Verdana" size="1" color="#4783C5"><%=getadminInvLngStr("DtxtBy")%></font><font face="Verdana" size="1" color="#4783C5"> </font><select size="1" name="MinInvBy" class="input">
			    <option value="W" <% If myApp.MinInvBy = "W" Then Response.Write "selected" %>><%=getadminInvLngStr("DtxtWarehouse")%></option>
				<option value="S" <% If myApp.MinInvBy = "S" Then Response.Write "selected" %>><%=getadminInvLngStr("DtxtSAP")%></option>
			    </select></td>
			</tr>
			<tr>
				<td width="300">
				<img src="images/ganchito.gif"><font face="Verdana" size="1" color="#4783C5">&nbsp;</font><input class="noborder" type="checkbox" name="EnableMinInvV" id="EnableMinInvV" value="Y" <% If myApp.EnableMinInvV Then %>checked<% End If %>><font face="Verdana" size="1" color="#4783C5"><label for="EnableMinInvV"><%=getadminInvLngStr("LtxtMinInvV")%></label></font></td>
				<td>
    			<input type="text" name="MinInvV" size="10" style="text-align:right" class="input" value="<%=FormatNumber(myApp.MinInvV,myApp.QtyDec)%>" onfocus="this.select()" onchange="chkThis(this, 0, <%=myApp.QtyDec%>, document.form1.oldMinInvV.value);document.form1.oldMinInvV.value=this.value;" onkeydown="return chkMax(event, this, 6);">
    			<input type="hidden" name="oldMinInvV" value="<%=FormatNumber(myApp.MinInvV,myApp.QtyDec)%>">
				<font face="Verdana" size="1" color="#4783C5"><%=getadminInvLngStr("DtxtBy")%> </font><select size="1" name="MinInvVBy" class="input">
			    <option value="W" <% If myApp.MinInvVBy = "W" Then Response.Write "selected" %>><%=getadminInvLngStr("DtxtWarehouse")%></option>
				<option value="S" <% If myApp.MinInvVBy = "S" Then Response.Write "selected" %>><%=getadminInvLngStr("DtxtSAP")%></option>
			    </select></td>
			</tr>
			<tr>
				<td width="300">
				<img src="images/ganchito.gif"><font face="Verdana" size="1" color="#4783C5">&nbsp;</font><input type="checkbox" name="ManageItmWhs" <% If myApp.ManageItmWhs Then %>checked<% End If %> id="ManageItmWhs" value="Y" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="ManageItmWhs"><%=getadminInvLngStr("LtxtManageItmWhs")%></label></font></td>
				<td>
				&nbsp;</td>
			</tr>
			<tr>
				<td colspan="2">
				<img src="images/ganchito.gif"><font face="Verdana" size="1" color="#4783C5">&nbsp;</font><input type="checkbox" name="EnableSearchItmSupp" <% If myApp.EnableSearchItmSupp Then %>checked<% End If %> id="EnableSearchItmSupp" value="Y" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="EnableSearchItmSupp"><%=getadminInvLngStr("LtxtEnableSearchItmSu")%></label></font></td>
			</tr>
			<tr>
				<td width="300">
				<img src="images/ganchito.gif"><font face="Verdana" size="1"><span style="font-size: 12px" 12px?? font-size:> </span>
				</font><font face="Verdana" size="1" color="#4783C5"><%=getadminInvLngStr("LtxtMinPrice")%></font></td>
				<td>
    			<input type="text" name="MinPrice" size="10" style="text-align:right" class="input" value="<%=FormatNumber(myApp.MinPrice,myApp.PriceDec)%>" onfocus="this.select()" onchange="chkThis(this, 0, <%=myApp.PriceDec%>, document.form1.oldMinPrice.value);document.form1.oldMinPrice.value=this.value;" maxlength="10">
    			<input type="hidden" name="oldMinPrice" value="<%=FormatNumber(myApp.MinPrice,myApp.PriceDec)%>"></td>
			</tr>
			<tr>
				<td width="300">
				<img src="images/ganchito.gif"><font face="Verdana" size="1" color="#4783C5">&nbsp;<%=getadminInvLngStr("Ltxtf_creacion")%></font></td>
				<td>
						<table cellpadding="0" cellspacing="0" border="0">
							<tr>
								<td><input type="text" name="f_creacion" id="f_creacion" size="10" style="text-align:right" class="input" value="<%=myApp.f_creacion%>" onfocus="this.select()" maxlength="10"></td>
								<td valign="middle">
								<table cellpadding="0" cellspacing="0" border="0">
									<tr>
										<td><img src="images/img_nud_up.gif" id="btnf_creacionUp"></td>
									</tr>
									<tr>
										<td><img src="images/spacer.gif"></td>
									</tr>
									<tr>
										<td><img src="images/img_nud_down.gif" id="btnf_creacionDown"></td>
									</tr>
								</table>
								</td>
							</tr>
						</table>
						<script language="javascript">NumUDAttach('form1', 'f_creacion', 'btnf_creacionUp', 'btnf_creacionDown');</script>
						</td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td bgcolor="#E1F3FD">&nbsp;<b><font face="Verdana" size="1" color="#31659C"><%=getadminInvLngStr("LttlCmb")%>
		</font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE"><font face="Verdana" size="1"> <img src="images/lentes.gif"> </font>
		<font face="Verdana" size="1" color="#4783C5"><%=getadminInvLngStr("LttlCmbNote")%></font></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table8" bgcolor="#F7FBFF">
			<tr>
				<td width="300">
				<img src="images/ganchito.gif"><font face="Verdana" size="1" color="#4783C5">&nbsp;</font><input type="checkbox" name="chkEnableCombos" <% If myApp.EnableCombos Then %>checked<% End If %> id="chkEnableCombos" value="Y" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="chkEnableCombos"><%=getadminInvLngStr("LtxtEnableCombos")%></label></font></td>
				<td>
				&nbsp;</td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td bgcolor="#E1F3FD">&nbsp;<b><font face="Verdana" size="1" color="#31659C"><%=getadminInvLngStr("LttlAvlMan")%>
		</font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE"><font face="Verdana" size="1"> <img src="images/lentes.gif"> </font>
		<font face="Verdana" size="1" color="#4783C5"><%=getadminInvLngStr("LttlAvlManNote")%></font></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table8" bgcolor="#F7FBFF">
			<tr>
				<td width="300">
				<p><img src="images/ganchito.gif"><font face="Verdana" size="1"><span style="font-size: 12px" 12px?? font-size:> </span>
				</font><font face="Verdana" size="1" color="#4783C5">
				<%=getadminInvLngStr("LtxtVerfyDisp")%></font></td>
				<td>
				<span style="font-size: 12px; font-size: 12px">
				<select size="1" name="VerfyDisp" class="input" onchange="javascript:document.form1.VerfyDispWhs.disabled = (this.value != 'S' && this.value != 'O')">
			    <option value="D" <% If myApp.VerfyDisp = "D" Then Response.Write "selected" %>><%=getadminInvLngStr("DtxtDisabled")%></option>
				<option value="S" <% If myApp.VerfyDisp = "S" Then Response.Write "selected" %>>
				<%=getadminInvLngStr("DtxtSAP")%></option>
				<option value="O" <% If myApp.VerfyDisp = "O" Then Response.Write "selected" %>>
				<%=getadminInvLngStr("DtxtOLK")%></option>
				<option value="C" <% If myApp.VerfyDisp = "C" Then Response.Write "selected" %>>
				<%=getadminInvLngStr("LtxtCustom")%></option>
			    </select>
			    </span>
			    </td>
			</tr>
			<tr>
				<td width="300"><img src="images/ganchito.gif"><font face="Verdana" size="1" color="#4783C5">
				<input type="checkbox" class="noborder" <% If myApp.VerfyDisp <> "S" and myApp.VerfyDisp <> "O" Then %>disabled<% End If %> name="VerfyDispWhs" <% If myApp.VerfyDispWhs = "S" Then %>checked<% End If %> id="VerfyDispWhs" value="S" style="width: 20px"><label for="VerfyDispWhs"><%=getadminInvLngStr("LtxtVerfyDispWhs")%></label></font></td>
				<td>
				</td>
			</tr>
			<% If 1 = 2 Then %>
			<tr>
				<td>
				<img src="images/ganchito.gif"><font face="Verdana" size="1" color="#4783C5">&nbsp;<%=getadminInvLngStr("LtxtVerfyDispMethod")%></font>
				</td>
				<td>
				<font face="Verdana" size="1" color="#4783C5">
				<input name="VerfyDispMethod" class="noborder" <% If myApp.VerfyDispMethod = "C" Then %>checked<% End If %> type="radio" value="C" id="VerfyDispMethodC"><label for="VerfyDispMethodC"><%=getadminInvLngStr("DtxtConfirm")%></label>
				<input name="VerfyDispMethod" class="noborder" <% If myApp.VerfyDispMethod = "M" Then %>checked<% End If %> type="radio" style="height: 20px" value="M" id="VerfyDispMethodM"><label for="VerfyDispMethodM"><%=getadminInvLngStr("DtxtError")%></label> </font></td>
			</tr>
			<% End IF %>

					</table>
		</td>
	</tr>
	<tr>
		<td bgcolor="#E1F3FD">&nbsp;<b><font face="Verdana" size="1" color="#31659C"><%=getadminInvLngStr("LttlCatFilter")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE"><font face="Verdana" size="1">
		<img src="images/lentes.gif">
		<font color="#4783C5"><%=getadminInvLngStr("LttlCatFilterNote")%></font></font></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<div align="left">
			<table border="0" cellpadding="0" width="100%" id="table9">
				<tr>
					<td width="300"><img src="images/ganchito.gif"><font face="Verdana" size="1"><span style="font-size: 12px" 12px?? font-size:> </span>
					</font><font face="Verdana" size="1" color="#4783C5">
					<%=getadminInvLngStr("LtxtApplyInvFiltersBy")%></font></td>
					<td>
					<span style="font-size: 12px">
					<select size="1" name="ApplyInvFiltersBy" class="input">
					<option <% If myApp.ApplyInvFiltersBy = "N" Then %>selected<% End If %> value="N">In</option>
					<option <% If myApp.ApplyInvFiltersBy = "I" Then %>selected<% End If %> value="I">Inner Join</option>
					</select></span></td>
				</tr>
				<tr>
					<td width="300" valign="top">
					<font face="Verdana" size="1" color="#4783C5"> 
					<%=getadminInvLngStr("DtxtQuery")%> - (ItemCode not in)</font></td>
					<td>
					<table cellpadding="0" cellspacing="0" border="0">
						<tr>
							<td rowspan="2">
								<textarea rows="10" name="GenFilter" dir="ltr" cols="87" class="input" onkeydown="javascript:document.form1.btnVerfyFilter.src='images/btnValidate.gif';document.form1.btnVerfyFilter.style.cursor = 'hand';;document.form1.valGenFilter.value='Y';"><% If Not IsNull(myApp.GenFilter) Then %><%=Server.HTMLEncode(myApp.GenFilter)%><% End If %></textarea>
							</td>
							<td valign="top">
								<img style="cursor: hand; " src="images/qry_note.gif" id="btnNoteFilter" alt="<%=getadminInvLngStr("DtxtDefinition")%>" onclick="javascript:doFldNote(2, 'GenFilter', -1, null);">
							</td>
						</tr>
						<tr>
							<td valign="bottom">
								<img src="images/btnValidateDis.gif" id="btnVerfyFilter" alt="<%=getadminInvLngStr("DtxtValidate")%>" onclick="javascript:if (document.form1.valGenFilter.value == 'Y')VerfyFilter();">	
								<input type="hidden" name="valGenFilter" value="N">
							</td>
						</tr>
					</table>
					</td>
				</tr>
				<tr>
					<td width="300" valign="top">
					<font size="1" color="#4783C5" face="Verdana"><%=getadminInvLngStr("LtxtAvlVars")%>:</font></td>
					<td>
					<font size="1" color="#4783C5" face="Verdana">
					<span dir="ltr">@CardCode</span> = <%=getadminInvLngStr("LtxtCCode")%><br>
					<span dir="ltr">@SlpCode</span> = <%=getadminInvLngStr("LtxtAgentCode")%><br>
					<span dir="ltr">@UserType</span> = <%=getadminInvLngStr("LtxtUserType")%></font></td>
				</tr>
				<tr>
					<td width="300" bgcolor="#F7FBFF">
					<img src="images/ganchito.gif"><font size="1" face="Verdana"> </font>
					<input type="checkbox" name="GenFAppV" <% If myApp.GenFAppV Then %>checked<% End If %> value="Y" id="GenFAppV" class="noborder"><font color="#4783C5" size="1" face="Verdana"><label for="GenFAppV"><%=getadminInvLngStr("LtxtGenFAppV")%></label></font></td>
					<td bgcolor="#F7FBFF">
					<p align="right">
					&nbsp;</td>
				</tr>
				<tr>
					<td width="300" bgcolor="#F7FBFF">
					<img src="images/ganchito.gif"><font size="1" face="Verdana">
					</font>
					<input type="checkbox" name="GenFAppC" <% If myApp.GenFAppC Then %>checked<% End If %> value="Y" id="GenFAppC" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="GenFAppC"><%=getadminInvLngStr("LtxtGenFAppC")%></label></font></td>
					<td bgcolor="#F7FBFF">
					<p align="right">
					&nbsp;</td>
				</tr>
			</table>
		</div>
		</td>
	</tr>
	<tr>
		<td bgcolor="#E1F3FD">&nbsp;<b><font face="Verdana" size="1" color="#31659C"><%=getadminInvLngStr("LttlItemRecQry")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE"><font face="Verdana" size="1">
		<img src="images/lentes.gif">
		<font color="#4783C5"><%=getadminInvLngStr("LttlItemRecQryNote")%></font></font></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<div align="left">
			<table border="0" cellpadding="0" width="100%">
				<tr>
					<td width="300" bgcolor="#F7FBFF">
					<img src="images/ganchito.gif"><font size="1" face="Verdana"> </font>
					<input type="checkbox" name="EnableItemRec" <% If myApp.EnableItemRec Then %>checked<% End If %> value="Y" id="EnableItemRec" class="noborder"><font color="#4783C5" size="1" face="Verdana"><label for="EnableItemRec"><%=getadminInvLngStr("LtxtEnableItemRec")%></label></font></td>
					<td bgcolor="#F7FBFF">
					<p align="right">
					&nbsp;</td>
				</tr>
				<tr>
					<td width="300" valign="top">
					<font face="Verdana" size="1" color="#4783C5"> 
					<%=getadminInvLngStr("LtxtItemRepQry")%>
					</font></td>
					<td>
					<table cellpadding="0" cellspacing="0" border="0">
						<tr>
							<td rowspan="2">
								<textarea rows="10" id="ItemRecQry" name="ItemRecQry" dir="ltr" cols="87" class="input" onkeydown="javascript:document.form1.btnVerfyItemRecFilter.src='images/btnValidate.gif';document.form1.btnVerfyItemRecFilter.style.cursor = 'hand';document.form1.valItemRecQry.value='Y';"><% If Not IsNull(myApp.ItemRecQry) Then %><%=Server.HTMLEncode(myApp.ItemRecQry)%><% End If %></textarea>
							</td>
							<td valign="top">
								<img src="images/qry_note.gif" id="btnItemRecFilter" alt="<%=getadminInvLngStr("DtxtDefinition")%>" onclick="javascript:doFldNote(2, 'ItemRec', -1, null);">
							</td>
						</tr>
						<tr>
							<td valign="bottom">
								<img src="images/btnValidateDis.gif" id="btnVerfyItemRecFilter" alt="<%=getadminInvLngStr("DtxtValidate")%>" onclick="javascript:if (document.form1.valItemRecQry.value == 'Y')VerfyItemRecFilter();">
								<input type="hidden" name="valItemRecQry" value="N">
							</td>
						</tr>
					</table>
					</td>
				</tr>
				<tr>
					<td width="300" valign="top">
					&nbsp;</td>
					<td class="style3">
					<font face="Verdana" size="1" color="#4783C5"><%=getadminInvLngStr("LtxtItemRepQryNote")%><span class="style5">*</span></font></td>
				</tr>
				<tr>
					<td width="300" valign="top">
					<font size="1" color="#4783C5" face="Verdana"><%=getadminInvLngStr("LtxtReqCols")%>:</font></td>
					<td class="style3">
					<font face="Verdana" size="1" color="#4783C5">
					ItemCode = <%=getadminInvLngStr("DtxtItemCode")%><br>
					Quantity = <%=getadminInvLngStr("DtxtQty")%><br>
					Locked = <%=getadminInvLngStr("DtxtLock")%> <br>
					Checked = <%=getadminInvLngStr("DtxtChecked")%> <br>
					WhsCode = <%=getadminInvLngStr("DtxtWhsCode")%> <br>
					Comment = <%=getadminInvLngStr("DtxtCommentaries")%> <br>
					</font></td>
				</tr>
				<tr>
					<td width="300" valign="top">
					<font size="1" color="#4783C5" face="Verdana"><%=getadminInvLngStr("LtxtAvlVars")%>:</font></td>
					<td class="style3">
					<font face="Verdana" size="1" color="#4783C5">
					<span dir="ltr">@ItemCode</span> = <%=getadminInvLngStr("DtxtItemCode")%><br>
					<span dir="ltr">@LogNum</span> = <%=getadminInvLngStr("DtxtLogNum")%> <br>
					<span dir="ltr">@CardCode</span> = <%=getadminInvLngStr("DtxtClientCode")%><br>
					<span dir="ltr">@SlpCode</span> = <%=getadminInvLngStr("DtxtAgentCode")%><br>
					<span dir="ltr">@branch</span> = <%=getadminInvLngStr("DtxtBranch")%><br>
					<span dir="ltr">@WhsCode</span> <%=getadminInvLngStr("DtxtWhsCode")%><br>
					<span dir="ltr">@LanID</span> = <%=getadminInvLngStr("DtxtLanID")%></font></td>
				</tr>
			</table>
		</div>
		</td>
	</tr>
	<tr>
		<td bgcolor="#E1F3FD">&nbsp;<b><font face="Verdana" size="1" color="#31659C"><%=getadminInvLngStr("LttlCodeBarQry")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE"><font face="Verdana" size="1">
		<img src="images/lentes.gif">
		<font color="#4783C5"><%=getadminInvLngStr("LttlCodeBarQryNote")%></font></font></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<div align="left">
			<table border="0" cellpadding="0" width="100%">
				<tr>
					<td width="300" bgcolor="#F7FBFF">
					<img src="images/ganchito.gif"><font size="1" face="Verdana"> </font>
					<input type="checkbox" name="EnableCodeBarsQry" <% If myApp.EnableCodeBarsQry Then %>checked<% End If %> value="Y" id="EnableCodeBarsQry" class="noborder"><font color="#4783C5" size="1" face="Verdana"><label for="EnableCodeBarsQry"><%=getadminInvLngStr("LtxtEnableCustomBarco")%></label></font></td>
					<td bgcolor="#F7FBFF">
					<p align="right">
					&nbsp;</td>
				</tr>
				<tr>
					<td width="300"><img src="images/ganchito.gif"><font face="Verdana" size="1"><span style="font-size: 12px"> </span>
					</font><font face="Verdana" size="1" color="#4783C5">
					<%=getadminInvLngStr("LtxtMethod")%></font></td>
					<td>
					<span style="font-size: 12px">
					<select size="1" name="CodeBarsQryMethod" class="input" onchange="changeMethod(this.value);">
					<option <% If myApp.CodeBarsQryMethod = "R" Then %>selected<% End If %> value="R">Replace</option>
					<option <% If myApp.CodeBarsQryMethod = "I" Then %>selected<% End If %> value="I">Inner Join</option>
					</select></span></td>
				</tr>

				<tr>
					<td width="300" valign="top">
					<font face="Verdana" size="1" color="#4783C5"> 
					<%=getadminInvLngStr("LtxtBarCodeQry")%><br><span id="txtCodeBarsSample"><%
					Select Case myApp.CodeBarsQryMethod
						Case "R" %>set @CodeBars = ([<%=getadminInvLngStr("DtxtQuery")%>])<%
						Case "I" %>select CodeBars from ([<%=getadminInvLngStr("DtxtQuery")%>])<%
					End Select %></span></font></td>
					<td>
					<table cellpadding="0" cellspacing="0" border="0">
						<tr>
							<td rowspan="2">
								<textarea rows="10" id="CodeBarsQry" name="CodeBarsQry" dir="ltr" cols="87" class="input" onkeydown="javascript:document.form1.btnVerfyBarCodeFilter.src='images/btnValidate.gif';document.form1.btnVerfyBarCodeFilter.style.cursor = 'hand';;document.form1.valCodeBarsQry.value='Y';"><% If Not IsNull(myApp.CodeBarsQry) Then %><%=Server.HTMLEncode(myApp.CodeBarsQry)%><% End If %></textarea>
							</td>
							<td valign="top">
								<img src="images/qry_note.gif" id="btnNoteFilter" alt="<%=getadminInvLngStr("DtxtDefinition")%>" onclick="javascript:doFldNote(2, 'CodeBarsQry', -1, null);">
							</td>
						</tr>
						<tr>
							<td valign="bottom">
								<img src="images/btnValidateDis.gif" id="btnVerfyBarCodeFilter" alt="<%=getadminInvLngStr("DtxtValidate")%>" onclick="javascript:if (document.form1.valCodeBarsQry.value == 'Y')VerfyBarCodeFilter();">
								<input type="hidden" name="valCodeBarsQry" value="N">
							</td>
						</tr>
					</table>
					</td>
				</tr>
				<tr>
					<td width="300" valign="top">
					<font size="1" color="#4783C5" face="Verdana"><%=getadminInvLngStr("LtxtAvlVars")%>:</font></td>
					<td>
					<font size="1" color="#4783C5" face="Verdana">
					<span dir="ltr">@CodeBars</span> = <%=getadminInvLngStr("LtxtBarCodeVar")%></font></td>
				</tr>
			</table>
		</div>
		</td>
	</tr>
	<% 
	set rSeries = Server.CreateObject("ADODB.RecordSet")
	set rd = Server.CreateObject("ADODB.RecordSet")
	sql = "select " & _
			"(select AlterDesc from OLKCommon..OLKAlterNames T0 where T0.LanID = " & Session("LanID") & " and T0.AlterID = 13) Desc14,  " & _
			"(select AlterDesc from OLKCommon..OLKAlterNames T0 where T0.LanID = " & Session("LanID") & " and T0.AlterID = 11) Desc15,  " & _
			"(select AlterDesc from OLKCommon..OLKAlterNames T0 where T0.LanID = " & Session("LanID") & " and T0.AlterID = 12) Desc16,  " & _
			"(select AlterDesc from OLKCommon..OLKAlterNames T0 where T0.LanID = " & Session("LanID") & " and T0.AlterID = 18) Desc19,  " & _
			"(select AlterDesc from OLKCommon..OLKAlterNames T0 where T0.LanID = " & Session("LanID") & " and T0.AlterID = 15) Desc20,  " & _
			"(select AlterDesc from OLKCommon..OLKAlterNames T0 where T0.LanID = " & Session("LanID") & " and T0.AlterID = 16) Desc21 "
	set rd = conn.execute(sql)
	
	sql = 		"select T0.ObjectCode, T0.Type, T0.DocType, T0.ChkShowReqSum, T0.ChkAllowOVerload, T0.ChkOp, IsNull(ChkOpSeries, -1) ChkOpSeries, ChkImpExp, ChkSerial, " & _
				"Case T0.ObjectCode " & _
				"	When 13 Then (select Plural from OLKCommon..OLKAlterNames where LanID = " & Session("LanID") & " and AlterID = 5) " & _
				"	When 15 Then (select Plural from OLKCommon..OLKAlterNames where LanID = " & Session("LanID") & " and AlterID = 11) " & _
				"	When 16 Then (select Plural from OLKCommon..OLKAlterNames where LanID = " & Session("LanID") & " and AlterID = 12) " & _
				"	When 17 Then (select Plural from OLKCommon..OLKAlterNames where LanID = " & Session("LanID") & " and AlterID = 7) " & _
				"	When 18 Then (select Plural from OLKCommon..OLKAlterNames where LanID = " & Session("LanID") & " and AlterID = 17) " & _
				"	When 20 Then (select Plural from OLKCommon..OLKAlterNames where LanID = " & Session("LanID") & " and AlterID = 15) " & _
				"	When 21 Then (select Plural from OLKCommon..OLKAlterNames where LanID = " & Session("LanID") & " and AlterID = 16) " & _
				"	When 22 Then (select Plural from OLKCommon..OLKAlterNames where LanID = " & Session("LanID") & " and AlterID = 14) " & _
				"End ObjectDesc " & _
				"from OLKInOutSettings T0 "
	rs.open sql, conn, 3, 1 %>
	<tr>
		<td bgcolor="#E1F3FD">&nbsp;<b><font face="Verdana" size="1" color="#31659C"><%=getadminInvLngStr("LtxtPurOrderCheck")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE"><font face="Verdana" size="1">
		<img src="images/lentes.gif">
		<font color="#4783C5"><%=getadminInvLngStr("LttlPurOrderNote")%></font></font></td>
	</tr>
	<tr>
		<td>
		<% doInOut "I" %>
		</td>
	</tr>
	<tr>
		<td bgcolor="#E1F3FD">&nbsp;<b><font face="Verdana" size="1" color="#31659C"><%=getadminInvLngStr("LtxtSalesOrderCheck")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE"><font face="Verdana" size="1">
		<img src="images/lentes.gif">
		<font color="#4783C5"><%=getadminInvLngStr("LttlSalesOrderNote")%></font></font></td>
	</tr>
	<tr>
		<td>
		<% doInOut "O" %>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table5">
			<tr>
				<td width="77">
				<input type="submit" value="<%=getadminInvLngStr("DtxtSave")%>" name="B1" class="OlkBtn"></td>
				<td><hr color="#0D85C6" size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
</table>
<input type="hidden" name="submitCmd" value="admininv">
</form>
<script language="javascript">
var objSeries;
function getSeries(o, ObjectCode)
{
	for (var i = o.length-1;i>=0;i--)
	{
		o.remove(i);
	}
	
	if (ObjectCode == -1 || ObjectCode == 17)
	{
		o.options[0] = new Option('<%=getadminInvLngStr("DtxtNotApply")%>', '');
		o.disabled = true;
	}
	else
	{
		objSeries = o;
		document.frmVerfyQuery.type.value = 'GetSeries';
		document.frmVerfyQuery.Query.value = ObjectCode;
		document.frmVerfyQuery.submit();
	}
}
function getObjSeries() { return objSeries; }

var verfyButton;
var hdverfyButton;
function VerfyBarCodeFilter()
{
	verfyButton = document.form1.btnVerfyBarCodeFilter;
	hdverfyButton = document.form1.valCodeBarsQry;
	document.frmVerfyQuery.type.value = 'CustomBarCode';
	document.frmVerfyQuery.Query.value = document.form1.CodeBarsQry.value;
	document.frmVerfyQuery.CodeBarsQryMethod.value = document.form1.CodeBarsQryMethod.value;
	if (document.frmVerfyQuery.Query.value != '')
	{
		document.frmVerfyQuery.submit();
	}
	else
	{
		VerfyBarCodeQueryVerified();
	}
}
function VerfyItemRecFilter()
{
	verfyButton = document.form1.btnVerfyItemRecFilter;
	hdverfyButton = document.form1.valItemRecQry;
	document.frmVerfyQuery.type.value = 'ItemRec';
	document.frmVerfyQuery.Query.value = document.form1.ItemRecQry.value;
	if (document.frmVerfyQuery.Query.value != '')
	{
		document.frmVerfyQuery.submit();
	}
	else
	{
		VerfyBarCodeQueryVerified();
	}
}
function VerfyFilter()
{
	verfyButton = document.form1.btnVerfyFilter;
	hdverfyButton = document.form1.valGenFilter;
	document.frmVerfyQuery.type.value = 'GenFilter';
	document.frmVerfyQuery.Query.value = document.form1.GenFilter.value;
	if (document.frmVerfyQuery.Query.value != '')
	{
		document.frmVerfyQuery.submit();
	}
	else
	{
		VerfyQueryVerified();
	}
}
function VerfyQueryVerified()
{
	//verfyButton.disabled = true;
	verfyButton.src='images/btnValidateDis.gif'
	verfyButton.style.cursor = '';
	hdverfyButton.value='N';
}
function changeMethod(value)
{
	javascript:document.form1.btnVerfyBarCodeFilter.src='images/btnValidate.gif';
	document.form1.btnVerfyBarCodeFilter.style.cursor = 'hand';
	document.form1.valCodeBarsQry.value='Y';
	
	var strSample = '';
	
	switch (value)
	{
		case 'R':
			strSample = 'set @CodeBars = ([<%=getadminInvLngStr("DtxtQuery")%>])';
			break;
		case 'I':
			strSample = 'select CodeBars from ([<%=getadminInvLngStr("DtxtQuery")%>])';
			break;
	}
	
	document.getElementById('txtCodeBarsSample').innerText = strSample;
}
//-->
</script>
<form name="frmVerfyQuery" action="verfyQuery.asp" method="post" target="iVerfyQuery">
	<iframe id="iVerfyQuery" name="iVerfyQuery" style="display: none" src=""></iframe>
	<input type="hidden" name="type" value="">
	<input type="hidden" name="Query" value="">
	<input type="hidden" name="parent" value="Y">
	<input type="hidden" name="CodeBarsQryMethod" value="">
</form>
<% Sub doInOut(ByVal t) %>
<table border="0" cellpadding="2" cellspacing="2" width="100%">
	<tr style="font-family: Verdana; font-size: xx-small; ">
		<td bgcolor="#E1F3FD" class="style2"><strong><%=getadminInvLngStr("LtxtSourceDoc")%></strong></td>
		<td class="style1"><strong><%=getadminInvLngStr("LtxtDocType")%></strong></td>
		<td class="style1"><strong><%=getadminInvLngStr("LtxtSaleChkShowReqSum")%></strong></td>
		<td class="style1"><strong><%=getadminInvLngStr("LtxtSaleChkAllowOverl")%></strong></td>
		<td class="style1"><strong><%=getadminInvLngStr("LtxtSaleChkOp")%></strong></td>
		<td class="style1"><strong><%=getadminInvLngStr("DtxtSeries")%></strong></td>
		<td class="style1"><strong><%=getadminInvLngStr("LtxtSaleChkImpExp")%></strong></td>
		<td class="style1"><strong><%=getadminInvLngStr("LtxtChkValSeries")%></strong></td>
	</tr>
	<% rs.Filter = "Type = '" & t & "'"
	rs.movefirst
	do while not rs.eof
	myID = rs("ObjectCode") & rs("Type") %>
	<tr style="font-family: Verdana; font-size: xx-small; ">
		<td bgcolor="#F7FBFF" class="style3"><%=rs("ObjectDesc")%><% If rs("ObjectCode") = 18 and rs("Type") = "I" or rs("ObjectCode") = 13 and rs("Type") = "O" Then %>&nbsp;(<%=getadminInvLngStr("DtxtReserved")%>)<% End If %></td>
		<td bgcolor="#F7FBFF" class="style3"><% Select Case rs("DocType")
		Case "S" %><%=getadminInvLngStr("LtxtSale")%>
		<% Case "P" %><%=getadminInvLngStr("LtxtPurchase")%>
		<% End Select %></td>
		<td bgcolor="#F7FBFF" align="center"><input type="checkbox" <% If rs("ChkShowReqSum") = "Y" Then %>checked<% End If %> name="ChkShowReqSum<%=myID%>" value="Y" class="noborder"></td>
		<td bgcolor="#F7FBFF" align="center"><input type="checkbox" <% If rs("ChkAllowOverload") = "Y" Then %>checked<% End If %> <% If rs("ObjectCode") <> 17 and rs("ObjectCode") <> 22 Then %>disabled<% End If %> name="ChkAllowOverl<%=myID%>" value="Y" class="noborder"></td>
		<td bgcolor="#F7FBFF">
		<select name="ChkOp<%=myID%>" size="1" class="input" onchange="javascript:getSeries(document.form1.Series<%=myID%>, this.value);">
		<option value="-1"><%=getadminInvLngStr("DtxtCheckOnly")%></option>
		<% If rs("ObjectCode") = 13 and rs("Type") = "I" Then %><option <% If rs("ChkOp") = 14 Then %>selected<% End If %> value="14"><%=rd("Desc14")%></option><% End If %>
		<% If rs("ObjectCode") = 15 Then %><option <% If rs("ChkOp") = 16 Then %>selected<% End If %> value="16"><%=rd("Desc16")%></option><% End If %>
		<% If rs("ObjectCode") = 16 Then %><option <% If rs("ChkOp") = 14 Then %>selected<% End If %> value="14"><%=rd("Desc14")%></option><% End If %>
		<% If rs("ObjectCode") = 17 Then %><option <% If rs("ChkOp") = 17 Then %>selected<% End If %> value="17"><%=getadminInvLngStr("LtxtUpdSalesOrder")%></option><% End If %>
		<% If rs("ObjectCode") = 17 or rs("ObjectCode") = 13 and rs("Type") = "O" Then %><option <% If rs("ChkOp") = 15 Then %>selected<% End If %> value="15"><%=rd("Desc15")%></option><% End If %>
		<% If rs("ObjectCode") = 18 and rs("Type") = "O" Then %><option <% If rs("ChkOp") = 19 Then %>selected<% End If %> value="19"><%=rd("Desc19")%></option><% End If %>
		<% If rs("ObjectCode") = 18 and rs("Type") = "I" or rs("ObjectCode") = 22 Then %><option <% If rs("ChkOp") = 20 Then %>selected<% End If %> value="20"><%=rd("Desc20")%></option><% End If %>
		<% If rs("ObjectCode") = 20 Then %><option <% If rs("ChkOp") = 21 Then %>selected<% End If %> value="21"><%=rd("Desc21")%></option><% End If %>
		<% If rs("ObjectCode") = 21 Then %><option <% If rs("ChkOp") = 19 Then %>selected<% End If %> value="19"><%=rd("Desc19")%></option><% End If %>
		</select></td>
		<td bgcolor="#F7FBFF">
		<select name="Series<%=myID%>" <% If rs("ChkOp") = -1 or rs("ChkOp") = 17 Then %>disabled<% End If %> size="1" class="input">
		<% If rs("ChkOp") = -1 or rs("ChkOp") = 17 Then %><option value=""><%=getadminInvLngStr("DtxtNotApply")%></option><% End If %>
		<%  If rs("ChkOp") <> -1 and rs("ChkOp") <> 17 Then
		GetQuery rSeries, 4, rs("ChkOp"), null
		do while not rSeries.eof %>
		<option <% If CInt(rSeries("Series")) = CInt(rs("ChkOpSeries")) Then %>selected<%end if%> value="<%=rSeries("Series")%>"><%=myHTMLEncode(rSeries("SeriesName"))%></option>
		<% rSeries.movenext
		loop
		End If %>
		</select></td>
		<td bgcolor="#F7FBFF" align="center"><input type="checkbox" <% If rs("ChkImpExp") = "Y" Then %>checked<% End If %> <% If rs("ObjectCode") <> 17 and rs("ObjectCode") <> 22 Then %>disabled<% End If %> name="ChkImpExp<%=myID%>" value="Y" class="noborder"></td>
		<td bgcolor="#F7FBFF">
		<select name="ChkSerial<%=myID%>" size="1" class="input">
		<option <% If rs("ChkSerial") = "N" Then %>selected<% End If %> value="N"><%=getadminInvLngStr("DtxtNone")%></option>
		<option <% If rs("ChkSerial") = "C" Then %>selected<% End If %> value="C"><%=getadminInvLngStr("DtxtConfirm")%></option>
		<option <% If rs("ChkSerial") = "E" Then %>selected<% End If %> value="E"><%=getadminInvLngStr("DtxtError")%></option>
		</select></td>
	</tr>
	<% rs.movenext
	loop %>
</table>
<% End Sub %><!--#include file="bottom.asp" -->
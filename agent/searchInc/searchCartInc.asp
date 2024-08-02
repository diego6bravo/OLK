<% addLngPathStr = "searchInc/" %>
<!--#include file="lang/searchCartInc.asp" -->

<head>
<style>
<!--
.input		
{

	
	color : #3366CC;
	font-family : Verdana, Arial, Helvetica, sans-serif;
	font-size : 10px;
	background-image: url('../menybg.gif');
	background-repeat: repeat-x;
	border: 1px solid #555555
}
.txtBox
{
border:1px solid #FFFFFF;
font-family: Verdana; 
font-size: 10px; 
color:#FFFFFF;
background-repeat: repeat-x;
background-color:#0065CE
}
-->
</style>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>
<table cellpadding="0" cellspacing="0" width="93%">
	<form method="POST" action="<% If Request("document") <> "B" Then %>search<% Else %>cart<% End If %>.asp" name="frmSmallSearch" onsubmit="return valSmallSearch();">
		<tr>
			<td>
			<p align="center"><b>
			<font size="1" face="Verdana" color="#FFFFFF"><%=myHTMLDecode(getsearchCartIncLngStr("LtxtSearch"))%>:</font></b></td>
		</tr>
		<tr>
			<td align="center">
			<input class="txtBox" type="text" name="string" size="16" value="<% If Request("string") <> "" Then Response.Write Server.HTMLEncode(Request("string"))%>" accesskey="<% If Session("myLng") = "es" or Session("myLng") = "pt" Then %>B<% ElseIf Session("myLng") = "he" Then %>&#1495;<% Else %>S<% End If %>" onfocus="this.select();focusSmallSearch(true);" onblur="focusSmallSearch(false);">
			</td>
		</tr>
		<% If myApp.SearchExactA Then %>
		<tr>
			<td>
			<p align="center">
			<font face="Verdana" size="1" color="#FFFFFF">
			<input type="radio" value="E" name="rdSearchAs" class="noborder" id="rdSearchAsE" <% If Request("rdSearchAs") = "" and myApp.SearchMethodA = "E" or Request("rdSearchAs") = "E" Then %>checked<% End If %>><label for="rdSearchAsE"><%=getsearchCartIncLngStr("DtxtExact")%></label>
			<input type="radio" name="rdSearchAs" class="noborder" id="rdSearchAsS" value="S" <% If Request("rdSearchAs") = "" and myApp.SearchMethodA = "L" or Request("rdSearchAs") = "S" Then %>checked<% End If %>><label for="rdSearchAsS"><%=getsearchCartIncLngStr("DtxtLike")%></label></font>
			</td>
		</tr>
		<% Else %>
		<input type="hidden" name="rdSearchAs" value="S">
		<% End If %>
		<tr>
			<td align="center">
			<input type="submit" value="<%=getsearchCartIncLngStr("LtxtGo")%>" name="B1" style="color: #FFFFFF; font-family: Verdana; font-size: 7pt; border: 1px solid #FFFFFF; background-color: #0066CB; " onfocus="focusSmallSearch(true);" onblur="focusSmallSearch(false);">
			</td>
		</tr>
		<tr>
			<td>
			<p align="center">
			<% ObjID = 4 %>
			<!--#include file="adCustomSearchInc.asp"-->
			</td>
		</tr>
      	<input type="hidden" name="cmd" value="<% If Request("document") <> "B" Then %>search<% If Session("Cart") = "cashCart" then response.write "cash" %>Cart<% Else %>cart<% End If %>">
      	<input type="hidden" name="orden1" value="<% If myApp.GetDefCatOrdr = "C" Then %>OITM.itemcode<% Else %>ItemName<% End If %>">
		<input type="hidden" name="orden2" value="asc">
		<tr id="trSmallSearch" style="display: none;" align="center">
			<td>
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td>
						<%
						If Request("document") = "" Then document = myApp.GetDefView Else document = Request("document")
						set objViewType = New clsViewType
						objViewType.ID = "document"
						objViewType.Value = document
						objViewType.AlterColor = true
						objViewType.OnClick = "document.frmSmallSearch.action = 'search.asp';document.frmSmallSearch.cmd.value = 'searchCart';document.frmSmallSearch.string.focus();document.getElementById('tdViewTypeCart').style.borderColor = 'transparent';"
						objViewType.doViewType
						%>
						</td>
						<% If myApp.EnableCartSum Then %><td id="tdViewTypeCart" style="padding-right: 2px;padding-left: 2px; border-bottom: 2px solid <% If Request("document") = "B" Then %>white<% Else %>transparent<% End If %>; "><img src="images/searchCartIcon.gif" alt="<%=getsearchCartIncLngStr("DtxtCart")%>" onclick="goSearchCart();"></td><% End If %>
					</tr>
				</table>
			</td>
		</tr>
		<input type="hidden" name="focus" value="frmSmallSearch.string">
	</form>
</table>
<script type="text/javascript">
var txtValEnterValue = '<%=getsearchCartIncLngStr("DtxtValEnterValue")%>';
var txtAddItmVal = '<%=getsearchCartIncLngStr("LtxtAddItmVal")%>';
var viewTypeCount = <%=viewTypeCount%>;

function chkSmallAddItm(field, fldUndo)
{
	if (field.value == '')
	{
		Field.value = fldUndo.value;
	}
	else if (!MyIsNumeric(field.value))
	{
		field.value = fldUndo.value;
	}
	else if (parseFloat(field.value) < 0)
	{
		field.value = fldUndo.value;
	}
	else if (parseFloat(field.value) > 99999999.999999)
	{
		field.value = 99999999.999999;
	}
	fldUndo.value = field.value;
}
</script>
<script type="text/javascript" src="searchInc/searchCartInc.js"></script>
<table cellpadding="0" cellspacing="0" border="0" width="100%">
	<tr>
		<td style="font-size: 4px;">&nbsp;</td>
	</tr>
</table>
<table cellpadding="0" cellspacing="0" width="93%">
	<form method="POST" action="cart/addcartsubmitm.asp" name="frmFastAdd">
	<input type="hidden" name="DocConf" value="">
	<tr>
		<td>
		<p align="center"><b>
		<font size="1" face="Verdana" color="#FFFFFF"><%=myHTMLDecode(getsearchCartIncLngStr("LtxtAdd"))%>:</font></b></td>
	</tr>
	<tr>
		<td>
		<p align="center"><b>
		<font size="1" face="Verdana" color="#FFFFFF"><%=getsearchCartIncLngStr("DtxtCode")%></font></b></td>
	</tr>
	<tr>
		<td>
		<input class="txtBox" type="text" name="Item" size="16" onfocus="this.select();focusSmallAddItm(true);" onblur="focusSmallAddItm(false);" onkeydown="chkExecCarInvAdd(event);" accesskey="<% If Session("rtl") = "" Then %>A<% Else %>&#1492;<% End If %>"></td>
	</tr>
	<tr id="trSmallAddItm" style="display: none;">
		<td>
		<p align="center"><b>
		<font size="1" face="Verdana" color="#FFFFFF"><%=getsearchCartIncLngStr("DtxtQty")%></font></b></td>
	</tr>
	<tr id="trSmallAddItm" style="display: none;">
		<td>
		<input class="txtBox" type="text" name="T1" size="16" value="1" onfocus="this.select();focusSmallAddItm(true);" onblur="focusSmallAddItm(false);" onkeydown="chkExecCarInvAdd(event);" onchange="javascript:chkSmallAddItm(this, document.frmFastAdd.T1Undo);" style="text-align: right">
		<input type="hidden" name="T1Undo" id="T1Undo" value="1"></td>
	</tr>
	<tr id="trSmallAddItm" style="display: none;">
		<td>
		<p align="center"><b>
		<font size="1" face="Verdana" color="#FFFFFF"><%=getsearchCartIncLngStr("DtxtUnit")%></font></b></td>
	</tr>
	<tr id="trSmallAddItm" style="display: none;">
		<td>
		<%
		If myApp.FastAddUnRem and Session("CurSaleType") <> "" Then
			fastAdd = Session("CurSaleType")
		Else
			fastAdd = myApp.GetSaleUnit
		End If %>
		<select size="1" name="SaleType" class="txtBox" style="width: 101px;" onfocus="focusSmallAddItm(true);" onblur="focusSmallAddItm(false);">
		<option value="1" <% If fastAdd = 1 Then %>selected<% End If%>>
		<%=getsearchCartIncLngStr("DtxtBaseUnit")%></option>
		<option value="2" <% If fastAdd = 2 Then %>selected<% End If%>>
		<%=getsearchCartIncLngStr("DtxtSalUnit")%></option>
		<option value="3" <% If fastAdd = 3 Then %>selected<% End If%>>
		<%=getsearchCartIncLngStr("DtxtPackUnit")%></option>
		</select></td>
	</tr>
	<% If myAut.HasAuthorization(68) Then %>
	<tr id="trSmallAddItm" style="display: none;">
		<td>
		<p align="center"><b>
		<font size="1" face="Verdana" color="#FFFFFF"><%=getsearchCartIncLngStr("DtxtPrice")%></font></b></td>
	</tr>
	<tr id="trSmallAddItm" style="display: none;">
		<td>
		<input class="txtBox" type="text" name="precio" id="searchCartIncPrice" size="16" maxlength="13" onfocus="focusSmallAddItm(true);" onkeydown="chkExecCarInvAdd(event);" onblur="focusSmallAddItm(false);" onchange="javascript:chkSmallAddItm(this, document.frmFastAdd.precioUndo);" onfocus="this.select();focusSmallAddItm(true);" onblur="focusSmallAddItm(false);" style="text-align: right">
		<input type="hidden" name="precioUndo" id="precioUndo" value=""></td>
	</tr>
	<% End If %>
	<tr id="trSmallAddItm" style="display: none;">
		<td>
		<p align="center">
		<input type="button" value="<%=getsearchCartIncLngStr("DtxtAdd")%>" name="B4" style="color: #FFFFFF; font-family: Verdana; font-size: 7pt; border: 1px solid #FFFFFF; background-color: #0066CB; width:76" onfocus="focusSmallAddItm(true);" onblur="focusSmallAddItm(false);" onclick="cartInvAddItem();"></td>
	</tr>
    <input type="hidden" name="redir" value="<% If Session("cart") = "cashInv" then %>cashInv<% Else %>cart<% End If %>">
	<input type="hidden" name="focus" value="frmFastAdd.Item">
	<input type="hidden" name="orden1" value="<% If myApp.GetDefCatOrdr = "C" Then %>OITM.ItemCode<% Else %>ItemName<% End If %>">
	<input type="hidden" name="AddPath" value="../">
	<input type="hidden" name="fastAdd" value="Y">
	</form>
</table>


<!--#include file="top.asp" -->
<!--#include file="RTE_configuration/browser_page_encoding_inc.asp" -->
<!--#include file="lang/adminCatProp.asp" -->

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style type="text/css">
.style1 {
				font-family: Verdana;
				font-size: xx-small;
				color: #4783C5;
}
.style2 {
				font-family: Verdana;
}
.style3 {
				color: #4783C5;
}
</style>
</head>
<% If Session("style") = "nc" Then %>
<br>
<% End If %>
<form method="POST" action="adminsubmit.asp" name="frmAdminCatProp">
<table border="0" cellpadding="0" width="100%" id="table3">
	<tr>
		<td bgcolor="#E1F3FD">&nbsp;<b><font face="Verdana" size="1" color="#31659C"><%=getadminCatPropLngStr("LtxtCatProp")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1">
		</font><font face="Verdana" size="1" color="#4783C5"><%=getadminCatPropLngStr("LtxtCatPropNote")%></font></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
			<div align="center">
				<table border="0" cellpadding="0" width="100%">
				<tr>
					<td bgcolor="#F7FBFF" width="435">
					<img src="images/ganchito.gif"><font size="1" face="Verdana">
					</font>
					<input type="checkbox" <% If myApp.SearchExactA Then %>checked<% End If %> name="SearchExactA" value="Y" id="SearchExactA" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="SearchExactA"><%=getadminCatPropLngStr("LtxtSearchExactA")%></label></font></td>
					<td bgcolor="#F7FBFF">
					<p>
					&nbsp;</td>
				</tr>
				<tr>
					<td bgcolor="#F7FBFF" width="435">
					<img src="images/ganchito.gif"><font size="1" face="Verdana">
					<font color="#4783C5"><%=getadminCatPropLngStr("LtxtSearchMethodA")%></font></font></td>
					<td bgcolor="#F7FBFF">
					<p>
					<select size="1" name="SearchMethodA" class="input">
					<option value="E" <% If myApp.SearchMethodA = "E" Then %>selected<%end if%>>
					<%=getadminCatPropLngStr("DtxtExact")%></option>
					<option value="L" <% If myApp.SearchMethodA = "L" Then %>selected<%end if%>>
					<%=getadminCatPropLngStr("DtxtLike")%></option>
					</select></td>
				</tr>
				<tr>
					<td bgcolor="#F7FBFF" width="435">
					<img src="images/ganchito.gif"><font size="1" face="Verdana">
					</font>
					<input type="checkbox" <% If myApp.SearchExactC Then %>checked<% End If %> name="SearchExactC" value="Y" id="SearchExactC" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="SearchExactC"><%=getadminCatPropLngStr("LtxtSearchExactC")%></label></font></td>
					<td bgcolor="#F7FBFF">
					<p>
					&nbsp;</td>
				</tr>
				<tr>
					<td bgcolor="#F7FBFF" width="435">
					<img src="images/ganchito.gif"><font size="1" face="Verdana">
					<font color="#4783C5"><%=getadminCatPropLngStr("LtxtSearchMethodC")%></font></font></td>
					<td bgcolor="#F7FBFF">
					<p>
					<select size="1" name="SearchMethodC" class="input">
					<option value="E" <% If myApp.SearchMethodC = "E" Then %>selected<%end if%>>
					<%=getadminCatPropLngStr("DtxtExact")%></option>
					<option value="L" <% If myApp.SearchMethodC = "L" Then %>selected<%end if%>>
					<%=getadminCatPropLngStr("DtxtLike")%></option>
					</select></td>
				</tr>
				<tr>
					<td bgcolor="#F7FBFF" width="435">
					<img src="images/ganchito.gif"><font size="1" face="Verdana">
					</font>
					<input type="checkbox" <% If myApp.SearchExactP Then %>checked<% End If %> name="SearchExactP" value="Y" id="SearchExactP" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="SearchExactP"><%=getadminCatPropLngStr("LtxtSearchExactP")%></label></font></td>
					<td bgcolor="#F7FBFF">
					<p>
					&nbsp;</td>
				</tr>
				<tr>
					<td bgcolor="#F7FBFF" width="435">
					<img src="images/ganchito.gif"><font size="1" face="Verdana">
					<font color="#4783C5"><%=getadminCatPropLngStr("LtxtSearchMethodP")%></font></font></td>
					<td bgcolor="#F7FBFF">
					<p>
					<select size="1" name="SearchMethodP" class="input">
					<option value="E" <% If myApp.SearchMethodP = "E" Then %>selected<%end if%>>
					<%=getadminCatPropLngStr("DtxtExact")%></option>
					<option value="L" <% If myApp.SearchMethodP = "L" Then %>selected<%end if%>>
					<%=getadminCatPropLngStr("DtxtLike")%></option>
					</select></td>
				</tr>
				<tr>
					<td bgcolor="#F7FBFF" width="435">
					<img src="images/ganchito.gif"><font size="1" face="Verdana">
					<font color="#4783C5"><%=getadminCatPropLngStr("LtxtDefCCatOrdr")%></font></font></td>
					<td bgcolor="#F7FBFF">
					<p>
					<select size="1" name="DefCatOrdrC" class="input">
					<option value="C" <% If myApp.DefCatOrdrC = "C" Then %>selected<%end if%>>
					<%=getadminCatPropLngStr("DtxtCode")%></option>
					<option value="N" <% If myApp.DefCatOrdrC = "N" Then %>selected<%end if%>>
					<%=getadminCatPropLngStr("DtxtDescription")%></option>
					</select></td>
				</tr>
				<tr>
					<td bgcolor="#F7FBFF" width="435">
					<img src="images/ganchito.gif"><font size="1" face="Verdana">
					<font color="#4783C5"><%=getadminCatPropLngStr("LtxtDefVCatOrdr")%></font></font></td>
					<td bgcolor="#F7FBFF">
					<p>
					<select size="1" name="DefCatOrdrV" class="input">
					<option value="C" <% If myApp.DefCatOrdrV = "C" Then %>selected<%end if%>>
					<%=getadminCatPropLngStr("DtxtCode")%></option>
					<option value="N" <% If myApp.DefCatOrdrV = "N" Then %>selected<%end if%>>
					<%=getadminCatPropLngStr("DtxtDescription")%></option>
					</select></td>
				</tr>
				<tr>
					<td bgcolor="#F7FBFF" width="435">
					<img src="images/ganchito.gif"><font size="1" face="Verdana">
					<font color="#4783C5"><%=getadminCatPropLngStr("LtxtDefViewCl")%></font></font></td>
					<td bgcolor="#F7FBFF">
					<select size="1" name="DefViewCL" class="input">
					<option value="T" <% If myApp.DefViewCL = "T" Then %>selected<%end if%>><%=getadminCatPropLngStr("DtxtStore")%></option>
					<option value="C" <% If myApp.DefViewCL = "C" Then %>selected<%end if%>><%=getadminCatPropLngStr("DtxtCat")%></option>
					<option value="L" <% If myApp.DefViewCL = "L" Then %>selected<%end if%>><%=getadminCatPropLngStr("DtxtList")%></option>
					</select>
					</td>
				</tr>
				<tr>
					<td bgcolor="#F7FBFF" width="435">
					<img src="images/ganchito.gif"><font size="1" face="Verdana">
					<font color="#4783C5"><%=getadminCatPropLngStr("LtxtDefViewAg")%></font></font></td>
					<td bgcolor="#F7FBFF">
					<select size="1" name="DefViewAG" class="input">
					<option value="T" <% If myApp.DefViewAG = "T" Then %>selected<%end if%>><%=getadminCatPropLngStr("DtxtStore")%></option>
					<option value="C" <% If myApp.DefViewAG = "C" Then %>selected<%end if%>><%=getadminCatPropLngStr("DtxtCat")%></option>
					<option value="L" <% If myApp.DefViewAG = "L" Then %>selected<%end if%>><%=getadminCatPropLngStr("DtxtList")%></option>
					</select>

					</td>
				</tr>
				<tr>
					<td bgcolor="#F7FBFF" width="435">
					<img src="images/ganchito.gif"><font size="1" face="Verdana">
					<font color="#4783C5"><%=getadminCatPropLngStr("LtxtAutoSearchOpen")%></font></font></td>
					<td bgcolor="#F7FBFF">
					<p>
					<select size="1" name="AutoSearchOpen" class="input">
					<option value="C" <% If myApp.AutoSearchOpen = "C" Then %>selected<%end if%>>
					<%=getadminCatPropLngStr("LoptAddToCart")%></option>
					<option value="N" <% If myApp.AutoSearchOpen = "N" Then %>selected<%end if%>>
					<%=getadminCatPropLngStr("LtxtRegCat")%></option>
					<option value="Y" <% If myApp.AutoSearchOpen = "Y" Then %>selected<%end if%>>
					<%=getadminCatPropLngStr("LtxtViewDet")%></option>
					</select></td>
				</tr>
				<tr>
					<td bgcolor="#F7FBFF" width="435">
					<img src="images/ganchito.gif"><font size="1" face="Verdana">
					</font>
					<input type="checkbox" <% If myApp.EnableMultCheck Then %>checked<% End If %> name="EnableMultCheck" value="Y" id="EnableMultCheck" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="EnableMultCheck"><%=getadminCatPropLngStr("LtxtEnableMultCheck")%></label></font></td>
					<td bgcolor="#F7FBFF">
					<p>
					&nbsp;</td>
				</tr>
				<tr>
					<td bgcolor="#F7FBFF" width="435">
					<img src="images/ganchito.gif"><font size="1" face="Verdana">
					</font>
					<input type="checkbox" <% If myApp.ShowClientRef Then %>checked<% End If %> name="ShowClientRef" value="Y" id="ShowClientRef" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="ShowClientRef"><%=getadminCatPropLngStr("LtxtShowClientRef")%></label></font></td>
					<td bgcolor="#F7FBFF">
					<p>
					&nbsp;</td>
				</tr>
				<% If 1 = 2 Then %>
				<tr>
					<td bgcolor="#F7FBFF" width="435">
					<img src="images/ganchito.gif"><font size="1" face="Verdana">
					</font>
					<input type="checkbox" <% If myApp.ShowNotAvlInv Then %>checked<% End If %> name="ShowNotAvlInv" value="Y" id="ShowNotAvlInv" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="ShowNotAvlInv">|L:txtShowNotAvlInv|</label></font></td>
					<td bgcolor="#F7FBFF">
					<p>
					&nbsp;</td>
				</tr>
				<% End If %>
				<tr>
					<td bgcolor="#F7FBFF" width="435">
					<img src="images/ganchito.gif"><font size="1" face="Verdana">
					</font>
					<input type="checkbox" <% If myApp.EnableOfertToDisc Then %>checked<% End If %> name="EnableOfertToDisc" value="Y" id="EnableOfertToDisc" class="noborder"><label for="EnableOfertToDisc"><font face="Verdana" size="1" color="#4783C5"><%=getadminCatPropLngStr("LtxtEnableOfertToDisc")%></font></label></td>
					<td bgcolor="#F7FBFF">
					<p>
					&nbsp;</td>
				</tr>
				<tr>
					<td bgcolor="#F7FBFF" width="435">
					<img src="images/ganchito.gif"><font size="1" face="Verdana">
					</font>
					<input type="checkbox" <% If myApp.EnableSearchAlterCode Then %>checked<% End If %> name="EnableSearchAlterCode" value="Y" id="EnableSearchAlterCode" class="noborder"><label for="EnableSearchAlterCode"><font face="Verdana" size="1" color="#4783C5"><%=getadminCatPropLngStr("LtxtEnableSearchAlter")%></font></label></td>
					<td bgcolor="#F7FBFF">
					<p>
					&nbsp;</td>
				</tr>
				<tr>
					<td bgcolor="#F7FBFF" width="435">
					<img src="images/ganchito.gif"><font size="1" face="Verdana">
					</font>
					<input type="checkbox" <% If myApp.SearchByVendorCode Then %>checked<% End If %> name="SearchByVendorCode" value="Y" id="SearchByVendorCode" class="noborder"><label for="SearchByVendorCode"><font face="Verdana" size="1" color="#4783C5"><%=getadminCatPropLngStr("LtxtSearchByVendorCod")%></font></label></td>
					<td bgcolor="#F7FBFF">
					<p>
					&nbsp;</td>
				</tr>
				<tr>
					<td bgcolor="#F7FBFF" width="435">
					<img src="images/ganchito.gif"><font size="1" face="Verdana">
					</font>
					<input type="checkbox" <% If myApp.ShowClientSalUn Then %>checked<% End If %> name="ShowClientSalUn" value="Y" id="ShowClientSalUn" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="ShowClientSalUn"><%=getadminCatPropLngStr("LtxtShowClientSalUn")%></label></font></td>
					<td bgcolor="#F7FBFF">
					<p>
					&nbsp;</td>
				</tr>
				<tr>
					<td bgcolor="#F7FBFF" width="435">
					<img src="images/ganchito.gif"><font size="1" face="Verdana">
					</font>
					<input type="checkbox" <% If myApp.ShowQtyInUnAg Then %>checked<% End If %> name="ShowQtyInUnAg" value="Y" id="ShowQtyInUnAg" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="ShowQtyInUnAg"><%=getadminCatPropLngStr("LtxtShowQtyInUnAg")%></label></font></td>
					<td bgcolor="#F7FBFF">
					<p>
					&nbsp;</td>
				</tr>
				<tr>
					<td bgcolor="#F7FBFF" width="435">
					<img src="images/ganchito.gif"><font size="1" face="Verdana">
					</font>
					<input type="checkbox" <% If myApp.ShowQtyInUnCl Then %>checked<% End If %> name="ShowQtyInUnCl" value="Y" id="ShowQtyInUnCl" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="ShowQtyInUnCl"><%=getadminCatPropLngStr("LtxtShowQtyInUnCl")%></label></font></td>
					<td bgcolor="#F7FBFF">
					<p>
					&nbsp;</td>
				</tr>
				<tr>
					<td bgcolor="#F7FBFF" width="435">
					<img src="images/ganchito.gif"><font size="1" face="Verdana">
					</font>
					<input type="checkbox" <% If myApp.ShowPocketImg Then %>checked<% End If %> name="ShowPocketImg" value="Y" id="ShowPocketImg" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="ShowPocketImg"><%=getadminCatPropLngStr("LtxtShowPocketImg")%></label></font></td>
					<td bgcolor="#F7FBFF">
					<p>
					&nbsp;</td>
				</tr>
				<tr>
					<td bgcolor="#F7FBFF" width="435">
					<img src="images/ganchito.gif"><font size="1" face="Verdana">
					</font>
					<input type="checkbox" <% If myApp.ShowClientImg Then %>checked<% End If %> name="ShowClientImg" value="Y" id="ShowClientImg" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="ShowClientImg"><%=getadminCatPropLngStr("LtxtShowClientImg")%></label></font></td>
					<td bgcolor="#F7FBFF">
					<p>
					&nbsp;</td>
				</tr>
				<tr>
					<td bgcolor="#F7FBFF" width="435">
					<img src="images/ganchito.gif"><font size="1" face="Verdana">
					</font>
					<input type="checkbox" <% If myApp.ShowAgentImg Then %>checked<% End If %> name="ShowAgentImg" value="Y" id="ShowAgentImg" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="ShowAgentImg"><%=getadminCatPropLngStr("LtxtShowAgentImg")%></label></font></td>
					<td bgcolor="#F7FBFF">
					<p>
					&nbsp;</td>
				</tr>
				<tr>
					<td bgcolor="#F7FBFF" width="435">
					<img src="images/ganchito.gif"><font size="1" face="Verdana">
					</font>
					<input type="checkbox" <% If myApp.ShowSearchTreeCount Then %>checked<% End If %> name="ShowSearchTreeCount" value="Y" id="ShowSearchTreeCount" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="ShowSearchTreeCount"><%=getadminCatPropLngStr("LShowSearchTreeCount")%></label></font></td>
					<td bgcolor="#F7FBFF">
					<p>
					&nbsp;</td>
				</tr>
				<tr>
					<td bgcolor="#F7FBFF" width="435">
					<img src="images/ganchito.gif"><font size="1" face="Verdana">
					</font>
					<input type="checkbox" <% If myApp.ShowSearchTreeSubCount Then %>checked<% End If %> name="ShowSearchTreeSubCount" value="Y" id="ShowSearchTreeSubCount" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="ShowSearchTreeSubCount"><%=getadminCatPropLngStr("LShowSearchTreeSubCou")%></label></font></td>
					<td bgcolor="#F7FBFF">
					<p>
					&nbsp;</td>
				</tr>
				<tr>
					<td bgcolor="#F7FBFF" width="435">
					<img src="images/ganchito.gif"><font size="1" face="Verdana">
					</font>
					<input type="checkbox" <% If myApp.EnableUnitSelection Then %>checked<% End If %> name="EnableUnitSelection" value="Y" id="EnableUnitSelection" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="EnableUnitSelection"><%=getadminCatPropLngStr("LEnableUnitSelection")%></label></font></td>
					<td bgcolor="#F7FBFF">
					<p>
					&nbsp;</td>
				</tr>
				<tr>
					<td bgcolor="#F7FBFF" width="435">
					<img src="images/ganchito.gif"><font size="1" face="Verdana">
					<font color="#4783C5"><%=getadminCatPropLngStr("LtxtAgentSaleUnit")%></font></font></td>
					<td bgcolor="#F7FBFF">
					<p>
					<select size="1" name="AgentSaleUnit" class="input">
					<option value="1" <% If myApp.AgentSaleUnit = "1" Then %>selected<%end if%>>
					<%=getadminCatPropLngStr("DtxtBaseUnit")%></option>
					<option value="2" <% If myApp.AgentSaleUnit = "2" Then %>selected<%end if%>>
					<%=getadminCatPropLngStr("DtxtSalUnit")%></option>
					<option value="3" <% If myApp.AgentSaleUnit = "3" Then %>selected<%end if%>>
					<%=getadminCatPropLngStr("DtxtPackUnit")%></option>
					</select></td>
				</tr>
				<tr>
					<td bgcolor="#F7FBFF" width="435">
					<img src="images/ganchito.gif"><font size="1" face="Verdana">
					<font color="#4783C5"><%=getadminCatPropLngStr("LtxtClientSaleUnit")%></font></font></td>
					<td bgcolor="#F7FBFF">
					<p>
					<select size="1" name="ClientSaleUnit" class="input">
					<option value="1" <% If myApp.ClientSaleUnit = "1" Then %>selected<%end if%>>
					<%=getadminCatPropLngStr("DtxtBaseUnit")%></option>
					<option value="2" <% If myApp.ClientSaleUnit = "2" Then %>selected<%end if%>>
					<%=getadminCatPropLngStr("DtxtSalUnit")%></option>
					<option value="3" <% If myApp.ClientSaleUnit = "3" Then %>selected<%end if%>>
					<%=getadminCatPropLngStr("DtxtPackUnit")%></option>
					</select></td>
				</tr>
				<tr>
					<td bgcolor="#F7FBFF" width="435">
					<img src="images/ganchito.gif"><font size="1" face="Verdana">
					</font>
					<input type="checkbox" <% If myApp.ShowPriceTax Then %>checked<% End If %> name="ShowPriceTax" value="Y" id="ShowPriceTax" class="noborder"><font face="Verdana" size="1" color="#4783C5"><label for="ShowPriceTax"><%=getadminCatPropLngStr("LtxtShowPriceTax")%></label></font></td>
					<td bgcolor="#F7FBFF">
					<p>
					&nbsp;</td>
				</tr>
				<tr>
					<td bgcolor="#F7FBFF" colspan="2">
					<img src="images/ganchito.gif"><font size="1" face="Verdana"> </font>
					<input type="checkbox" name="UnEmbPriceSet" value="Y" <% If myApp.UnEmbPriceSet Then %>checked<%end if %> id="UnEmbPriceSet" class="noborder"><label for="UnEmbPriceSet"><font face="Verdana" size="1" color="#4783C5"><%=getadminCatPropLngStr("LtxtUnEmbPriceSet")%></font></label></td>
				</tr>
				<tr>
					<td bgcolor="#F7FBFF" width="435">
					<img src="images/ganchito.gif"><font size="1" face="Verdana">
					<font color="#4783C5"><%=getadminCatPropLngStr("LtxtolkItemReport2")%></font></font></td>
					<td bgcolor="#F7FBFF">
					<p>
					<select size="1" name="olkItemReport2" class="input">
					<option <% If myApp.olkItemReport2 = "D" Then %>selected<% End If %> value="D">
					<%=getadminCatPropLngStr("DtxtDisabled")%></option>
					<option <% If myApp.olkItemReport2 = "L" Then %>selected<% End If %> value="L">
					<%=getadminCatPropLngStr("DtxtLow")%></option>
					<option <% If myApp.olkItemReport2 = "M" Then %>selected<% End If %> value="M">
					<%=getadminCatPropLngStr("DtxtMedium")%></option>
					<option <% If myApp.olkItemReport2 = "H" Then %>selected<% End If %> value="H">
					<%=getadminCatPropLngStr("DtxtHigh")%></option>
					</select></td>
				</tr>
				<tr>
					<td width="435" bgcolor="#F7FBFF">
					<img src="images/ganchito.gif"><font face="Verdana" size="1">
					<font color="#4783C5"><%=getadminCatPropLngStr("LtxtCarArt")%></font></font></td>
					<td bgcolor="#F7FBFF"><font face="Verdana" size="1">
					<select size="1" name="CarArt" class="input">
					<option value="-1">-- <%=getadminCatPropLngStr("DtxtDisabled")%> --</option>
					<% 
					GetQuery rd, 3, null, null
					do while not rd.eof %>
					<option <% If CInt(rd("ItmsTypCod")) = CInt(myApp.CarArt) Then Response.write "selected" %> value="<%=rd("ItmsTypCod")%>">
					<%=myHTMLEncode(rd("ItmsGrpNam"))%></option>
					<% 
					rd.movenext
					loop 
					%></select></font></td>
				</tr>
				<tr>
					<td width="435" bgcolor="#F7FBFF" valign="top">
					<img src="images/ganchito.gif"><font face="Verdana" size="1" color="#4783C5">
					<%=getadminCatPropLngStr("LtxtQryGroupSearch")%></font></td>
					<td bgcolor="#F7FBFF">
					<div id="listQryGroupSearch" style="width: 300px; height: 120px; background-color:#DEF7FF; overflow=auto; border-width: 1px; border-color: #6B6B6B;">
					<%
					check = True
					sql = 	"select T0.ItmsTypCod, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITG', 'ItmsGrpNam', T0.ItmsTypCod, T0.ItmsGrpNam) ItmsGrpNam, Case When T1.ItmsTypCod is not null Then 'Y' Else 'N' End Verfy " & _
							"from OITG T0 " & _
							"left outer join OLKSearchQryGroups T1 on T1.ItmsTypCod = T0.ItmsTypCod " & _
							"order by T0.ItmsTypCod"
					set rd = conn.execute(sql)
					do while not rd.eof
					If rd("Verfy") = "N" Then check = False %>
					<div><font face="Verdana" size="1" color="#4783C5"><input class="noborder" type="checkbox" name="chkQryGroup" id="chkQryGroup<%=rd(0)%>" value="<%=rd(0)%>" onclick="chkAllQryGroups();" <% If rd("Verfy") = "Y" Then %>checked<% End If %>><label for="chkQryGroup<%=rd(0)%>"><%=rd(1)%></label></font></div>
					<% rd.movenext
					loop %>
					</div>
					<div style="width: 300px; background-color:#DEF7FF;"><font face="Verdana" size="1" color="#4783C5"><input class="noborder" type="checkbox" <% If check Then %>checked<% End If %> name="chkQryGroupAll" id="chkQryGroupAll" value="Y" onclick="changeAllQryGroups(this.checked);"><label for="chkQryGroupAll"><%=getadminCatPropLngStr("DtxtAll")%></label></font></div>
					</td>
				</tr>
				</table>
			</div>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr>
				<td width="77">
				<input type="submit" value="<%=getadminCatPropLngStr("DtxtSave")%>" name="btnSave" class="OlkBtn"></td>
				<td><hr color="#0D85C6" size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
</table>
<input type="hidden" name="submitCmd" value="adminCatProp">
</form>
<script type="text/javascript">
function changeAllQryGroups(check)
{
	for (var i = 0;i<document.frmAdminCatProp.chkQryGroup.length;i++)
	{
		document.frmAdminCatProp.chkQryGroup[i].checked = check;
	}
}
function chkAllQryGroups()
{
	var check = true;
	for (var i = 0;i<document.frmAdminCatProp.chkQryGroup.length;i++)
	{
		if (!document.frmAdminCatProp.chkQryGroup[i].checked)
		{
			check = false;
		}
	}
	document.frmAdminCatProp.chkQryGroupAll.checked = check;
}
</script>
<!--#include file="bottom.asp" -->
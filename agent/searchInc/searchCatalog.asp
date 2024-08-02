<% addLngPathStr = "searchInc/" %>
<!--#include file="lang/searchCatalog.asp" -->


<% 
If Request.Form.Count > 0 Then
	If Request("CPList") <> "" Then Session("CPList") = Request("CPList") Else Session("CPList") = ""
End If
 %>
<script type="text/javascript">
function valSmallSearch()
{
	if (document.frmSmallSearch.sourceDoc.value != '' && (document.frmSmallSearch.DocNum.value == '' || !MyIsNumeric(document.frmSmallSearch.DocNum.value)))
	{
		alert("<%=getsearchCatalogLngStr("LtxtValDocNum")%>");
		document.frmSmallSearch.DocNum.focus();
		return false;
	}
	else if (document.frmSmallSearch.CPList.value == "X" && document.frmSmallSearch.sourceDoc.value == "")
	{
		alert("<%=getsearchCatalogLngStr("LtxtDocSrc")%>");
		return false;
	}
	return true;
}
</script>
<table border="0" cellpadding="0" cellspacing="0" width="93%">
			<form method="POST" action="search.asp" name="frmSmallSearch" onsubmit="return valSmallSearch();">
				<tr>
					<td>
					<p align="center"><b>
					<font size="1" face="Verdana" color="#FFFFFF"><%=myHTMLDecode(getsearchCatalogLngStr("LtxtSearch"))%></font></b></td>
				</tr>
				<tr>
					<td>
					<input class="input" type="text" name="string" size="18" style="border:1px solid #FFFFFF; font-family: Verdana; font-size: 10px; color:#FFFFFF; background-color:#0065CE" value="<% If Request("string") <> "" Then Response.Write Server.HTMLEncode(Request("string"))%>" onfocus="this.select()" accesskey="<% If Session("myLng") = "es" or Session("myLng") = "pt" Then %>B<% ElseIf Session("myLng") = "he" Then %>&#1495;<% Else %>S<% End If %>"></td>
				</tr>
				<% If myApp.SearchExactA Then %>
				<tr>
					<td>
					<p align="center">
					<font face="Verdana" size="1" color="#FFFFFF">
					<input type="radio" value="E" name="rdSearchAs" class="noborder" id="rdSearchAsE" <% If Request("rdSearchAs") = "" and myApp.SearchMethodA = "E" or Request("rdSearchAs") = "E" Then %>checked<% End If %>><label for="rdSearchAsE"><%=getsearchCatalogLngStr("DtxtExact")%></label>
					<input type="radio" name="rdSearchAs" class="noborder" id="rdSearchAsS" value="S" <% If Request("rdSearchAs") = "" and myApp.SearchMethodA = "L" or Request("rdSearchAs") = "S" Then %>checked<% End If %>><label for="rdSearchAsS"><%=getsearchCatalogLngStr("DtxtLike")%></label></font>
					</td>
				</tr>
				<% Else %>
				<input type="hidden" name="rdSearchAs" value="S">
				<% End If %>
				<tr>
					<td>
					<p align="center"><b>
					<font size="1" face="Verdana" color="#FFFFFF"><%=getsearchCatalogLngStr("LtxtPriceList")%></font></b></td>
				</tr>
				<tr>
					<td>
					<select size="1" name="CPList" style="border:1px solid #FFFFFF; font-family: Verdana; font-size: 10px; width: 104; background-color:#0065CE; color:#FFFFFF">
					<option value=""><%=getsearchCatalogLngStr("LtxtNoPrice")%></option>
					<option <% If Request("CPList") = "X" Then %>selected<% End If %> value="X"><%=getsearchCatalogLngStr("LtxtDocPrice")%></option>
					<% 
					set cmd = Server.CreateObject("ADODB.Command")
					cmd.ActiveConnection = connCommon
					cmd.CommandType = &H0004
					cmd.CommandText = "DBOLKGetPriceListFiltered" & Session("ID")
					cmd.Parameters.Refresh()
					cmd("@LanID") = Session("LanID")
					cmd("@UserAccess") = Session("UserAccess")
					cmd("@SlpCode") = Session("vendid")
					SET RS = cmd.execute()
					Do While NOT RS.EOF %>
					<option <% If CStr(Session("CPList")) = CStr(rs("ListNum")) then response.write "selected" %> value="<%=RS("Listnum")%>"><%=myHTMLEncode(RS("ListName"))%></option>
					<% RS.MoveNext
					loop %>
					</select></td>
				</tr>
				<tr>
					<td>
					<p align="center"><b>
					<font size="1" face="Verdana" color="#FFFFFF"><%=getsearchCatalogLngStr("LtxtDoc")%></font></b></td>
				</tr>
				<tr>
					<td>
					<select size="1" name="sourceDoc" style="border:1px solid #FFFFFF; font-family: Verdana; font-size: 10px; width: 104px; background-color:#0065CE; color:#FFFFFF; ">
					<option value=""><%=getsearchCatalogLngStr("DtxtCat")%></option>
					<optgroup label="<%=getsearchCatalogLngStr("LtxtSale")%>">
						<option value="23"<% If Request("sourceDoc") = "23" Then %> selected<% end if %>><% If 1 = 2 Then %>Cotizaci�n<% Else %><%=myHTMLEncode(txtQuote)%><% End If %></option>
						<option value="17"<% If Request("sourceDoc") = "17" Then %> selected<% end if %>><% If 1 = 2 Then %>Pedido<% Else %><%=myHTMLEncode(txtOrdr)%><% End If %></option>
						<option value="15"<% If Request("sourceDoc") = "15" Then %> selected<% end if %>><% If 1 = 2 Then %>Despacho<% Else %><%=myHTMLEncode(txtOdln)%><% End If %></option>
						<option value="13"<% If Request("sourceDoc") = "13" Then %> selected<% end if %>><% If 1 = 2 Then %>Factura<% Else %><%=myHTMLEncode(txtInv)%><% End If %></option>
					</optgroup>
					<optgroup label="<%=getsearchCatalogLngStr("LtxtPur")%>">
						<option value="22"<% If Request("sourceDoc") = "22" Then %> selected<% end if %>><% If 1 = 2 Then %>Orden de Compra<% Else %><%=myHTMLEncode(txtOpor)%><% End If %></option>
						<option value="20"<% If Request("sourceDoc") = "20" Then %> selected<% end if %>><% If 1 = 2 Then %>Entrada de Mercancia OP<% Else %><%=txtOpdn%><% End If %></option>
						<option value="18"<% If Request("sourceDoc") = "18" Then %> selected<% end if %>><% If 1 = 2 Then %>Comp. de Compra<% Else %><%=myHTMLEncode(txtOpch)%><% End If %></option>
					</optgroup>
					<optgroup label="<%=getsearchCatalogLngStr("DtxtOLK")%>">
						<option value="-4"<% If Request("sourceDoc") = "-4" Then %> selected<% End If %>><%=getsearchCatalogLngStr("DtxtLogNum")%></option>
					</optgroup>
					</select></td>
				</tr>
				<tr>
					<td>
					<p align="center"><b>
					<font size="1" face="Verdana" color="#FFFFFF"><%=getsearchCatalogLngStr("LtxtDocNum")%></font></b></td>
				</tr>
				<tr>
					<td>
					<input class="input" type="text" name="DocNum" size="18" style="border:1px solid #FFFFFF; font-family: Verdana; font-size: 10px; color:#FFFFFF; background-color:#0065CE" onfocus="this.select()" onkeydown="return valKeyNum(event);" value="<%=Request("DocNum")%>"></td>
				</tr>
				<tr>
					<td style="font-size: 4px">&nbsp;</td>
				</tr>
				<tr>
					<td>
					<p align="center">
									<input type="submit" value="<%=getsearchCatalogLngStr("DbtnSearch")%>" name="B1" style="color: #FFFFFF; font-family: Verdana; font-size: 7pt; border: 1px solid #FFFFFF; background-color: #0066CB; width:76"></td>
				</tr>
				<tr>
					<td style="font-size: 4px">&nbsp;</td>
				</tr>
				<tr>
					<td>
					<p align="center">
					<% ObjID = 4 %>
					<!--#include file="adCustomSearchInc.asp"-->
					</td>
				</tr>
              	<input type="hidden" name="cmd" value="searchCatalog">
				<input type="hidden" name="orden2" value="asc">
			<tr>
				<td align="center">
					<%
					If Request("document") = "" Then document = myApp.GetDefView Else document = Request("document")
					set objViewType = New clsViewType
					objViewType.ID = "document"
					objViewType.Value = document
					objViewType.AlterColor = true
					objViewType.doViewType
					%>
					</td>
			</tr>
				<input type="hidden" name="focus" value="frmSmallSearch.string">
			<input type="hidden" name="orden1" value="<% If myApp.GetDefCatOrdr = "C" Then %>OITM.ItemCode<% Else %>ItemName<% End If %>">
			</form>
		</table>
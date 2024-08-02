<!--#include file="clientInc.asp"-->
<!--#include file="agentTop.asp"-->
<% addLngPathStr = "" %>
<!--#include file="lang/ofertAgentContraOfert.asp" -->
<%
Dim fltParam
	For each item in Request("flt") 
		fltParam = fltParam & "&flt=" & item
	next

Dim errmsg
      set rs = Server.CreateObject("ADODB.recordset")

      sql = "select T0.ItemCode, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'ItemName', T0.ItemCode, ItemName) ItemName, " & _
			" OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'SalUnitMsr', T0.ItemCode, SalUnitMsr) SalUnitMsr, " & _
			" OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'SalPackMsr', T0.ItemCode, SalPackMsr) SalPackMsr, " & _
			" Replace(OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'UserText', T0.ItemCode, UserText) collate database_default, Char(13), '<br>') UserText, " & _
			"PicturName, BasePrice, ofertPrice, ofertQuantity, " & _
      	    "IsNull(OfertNote, '') OfertNote, ofertDiscount, IsNull(CardName, '') CardName, " & _
      	    "DateAdd(day, ofertLimit, responseDate) ofertLimit, " & _
      	    "responsePrice, responseQuantity, responseDiscount, responseLimit, IsNull(responseNote, '') responseNote, " & _
      		"(select ClientSaleUnit from olkcommon) SaleType, " & _
      		"NumInSale, SalPackUn " & _
      	    "from olkoferts T0 inner join oitm T1 on T1.ItemCode = T0.ItemCode " & _
      	    "inner join ocrd on ocrd.cardcode = T0.UserName " & _ 
      	    "inner join olkOfertsLines T2 on T2.ofertIndex = T0.ofertIndex " & _
      	    "where T0.OfertIndex = " & Request("ofertIndex") & " and ofertLineNum = " & _
      	    "(select top 1 ofertLineNum from olkOfertsLines where ofertIndex = T0.ofertIndex order by ofertLineNum desc)"
      		set rs = conn.execute(sql)	
		SaleType = rs("SaleType")
      UserName = rs("CardName")
		  If rs("PicturName") <> "" Then
		  Pic = rs("PicturName")'
		  Else
		  Pic = "n_a.gif"
		  End If 
		  
		  Select Case SaleType
		  	Case 1
		  		BasePrice = CDbl(rs("BasePrice"))
		  		OfertPrice = CDbl(rs("OfertPrice"))
		  		UnPrice = "Un."
		  		SaleUn = "Un.(1)"
		  		ofertQuantity = rs("ofertQuantity")
		  		responseQuantity = rs("responseQuantity")
		  		ofertTotal = CDbl(rs("ofertQuantity"))*CDbl(rs("ofertPrice"))
		  		responseTotal = CDbl(rs("responseQuantity"))*CDbl(rs("responsePrice"))
		  		responsePrice = CDbl(rs("responsePrice"))
		  	Case 2
		  		BasePrice = CDbl(rs("BasePrice"))*CDbl(rs("NumInSale"))
		  		OfertPrice = CDbl(rs("OfertPrice"))*CDbl(rs("NumInSale"))
		  		UnPrice = rs("SalUnitMsr")
		  		SaleUn = rs("SalUnitMsr") 
		  		If myApp.GetShowQtyInUn Then SaleUn = SaleUn & "(" & rs("NumInSale") & ")"
		  		ofertQuantity = rs("ofertQuantity")
		  		responseQuantity = rs("responseQuantity")
		  		ofertTotal = CDbl(rs("ofertQuantity"))*CDbl(rs("ofertPrice"))*CDbl(rs("NumInSale"))
		  		responseTotal = CDbl(rs("responseQuantity"))*CDbl(rs("responsePrice"))*CDbl(rs("NumInSale"))
		  		responsePrice = CDbl(rs("responsePrice"))*CDbl(rs("NumInSale"))
		  	Case 3
			  	SaleUn = rs("SalPackMsr")
			  	If myApp.GetShowQtyInUn Then SaleUn = SaleUn & "(" & rs("SalPackUn") & ")"
		  		If Not myApp.UnEmbPriceSet Then
		  			BasePrice = CDbl(rs("BasePrice"))*CDbl(rs("NumInSale"))
			  		OfertPrice = CDbl(rs("OfertPrice"))*CDbl(rs("NumInSale"))
			  		responsePrice = CDbl(rs("responsePrice"))*CDbl(rs("NumInSale"))
		  			UnPrice = rs("SalUnitMsr")
		  		ElseIf myApp.UnEmbPriceSet Then 
		  			BasePrice = CDbl(rs("BasePrice"))*CDbl(rs("NumInSale"))*CDbl(rs("SalPackUn"))
			  		OfertPrice = CDbl(rs("OfertPrice"))*CDbl(rs("NumInSale"))*CDbl(rs("SalPackUn"))
			  		responsePrice = CDbl(rs("responsePrice"))*CDbl(rs("NumInSale"))*CDbl(rs("SalPackUn"))
		  			UnPrice = rs("SalPackMsr")
			  		SaleUn = SaleUn & " x " & rs("SalUnitMsr") 
			  		If myApp.GetShowQtyInUn Then SaleUn = SaleUn & "(" & rs("NumInSale") & ")"
		  		End If
		  		ofertQuantity = CDbl(rs("ofertQuantity"))/CDbl(rs("SalPackUn"))
		  		responseQuantity = CDbl(rs("responseQuantity"))/CDbl(rs("SalPackUn"))
		  		ofertTotal = CDbl(rs("ofertQuantity"))*CDbl(rs("ofertPrice"))*CDbl(rs("NumInSale"))
		  		responseTotal = CDbl(rs("responseQuantity"))*CDbl(rs("responsePrice"))*CDbl(rs("NumInSale"))
		  End Select
%>
<script type="text/javascript">
function Start(page, w, h, s) {
OpenWin = this.open(page, "ImageThumb", "toolbar=no,menubar=no,location=no,scrollbars="+s+",resizable=yes, width="+w+",height="+h);
}

var SaleType = <%=SaleType%>;
var NumInSale = <%=rs("NumInSale")%>;
var SalPackUn = <%=rs("SalPackUn")%>;

function chkThis(Field)
{
	if (!MyIsNumeric(getNumericVB(Field.value)))
	{
		Field.value = '<%=BasePrice%>';
	}
	else if (parseFloat(getNumericVB(Field.value)) < 0)
	{
		Field.value = 0;
	}
	
	Field.value = OLKFormatNumber(parseFloat(getNumericVB(Field.value)),<%=myApp.PriceDec%>);
	
	if (parseFloat(getNumericVB(Field.value)) > parseFloat(getNumericVB("<%=BasePrice%>"))) Field.value = OLKFormatNumber(parseFloat(getNumericVB("<%=BasePrice%>")),<%=myApp.PriceDec%>);
	
	document.confirmar.ResponseDiscount.value = OLKFormatNumber(100-((100*parseFloat(getNumericVB(Field.value)))/parseFloat(getNumericVB("<%=BasePrice%>"))),<%=myApp.PriceDec%>);
}

function changeDiscount(Field)
{
	if (!MyIsNumeric(parseFloat(getNumericVB(Field.value))))
	{
		Field.value = "0";
	}
	else
	{
		document.confirmar.ResponsePrice.value = OLKFormatNumber(parseFloat(getNumericVB("<%=BasePrice%>"))-(parseFloat(getNumericVB("<%=BasePrice%>"))*parseFloat(getNumericVB(Field.value)))/100,<%=myApp.PriceDec%>);
	}
}

function chkNum(Field, Val)
{
	if (!MyIsNumeric(parseFloat(getNumericVB(Field.value))))
	{
		Field.value = Val;
	}
	else if (parseFloat(getNumericVB(Field.value)) <= 0)
	{
		Field.value = Val;
	}
}

function setTotal()
{
	if (document.confirmar.ResponseQuantity.value == '') document.confirmar.ResponseQuantity.value = 1;
	document.confirmar.total.value = OLKFormatNumber(parseFloat(getNumericVB(document.confirmar.ResponsePrice.value))*parseFloat(getNumericVB(document.confirmar.ResponseQuantity.value))<% If Not myApp.UnEmbPriceSet Then %>*SalPackUn<% End If %>,<%=myApp.PriceDec%>);
}

function changeTotal(Field)
{
	if (!MyIsNumeric(parseFloat(getNumericVB(Field.value))))
	{ 
		setTotal();
	}
	else if (parseFloat(getNumericVB(Field.value)) < 0)
	{
		Field.value = OLKFormatNumber(0, <%=myApp.PriceDec%>);
	}
	else
	{
		Field.value = OLKFormatNumber(parseFloat(getNumericVB(Field.value)),<%=myApp.PriceDec%>);
	}
	document.confirmar.ResponsePrice.value = OLKFormatNumber(parseFloat(getNumericVB(Field.value))/parseFloat(getNumericVB(document.confirmar.ResponseQuantity.value))<% If Not myApp.UnEmbPriceSet Then %>/SalPackUn<% End If %>,<%=myApp.PriceDec%>);
	chkThis(document.confirmar.ResponsePrice);
}

function FormatQty(Field, Dec)
{
	if (MyIsNumeric(getNumericVB(Field.value)))
	{
		if (parseFloat(getNumericVB(Field.value)) < 0) Field.value = 0;
		Field.value = OLKFormatNumber(parseFloat(getNumericVB(Field.value)),Dec);
	}
}

function doBlink() 
	{
	var blink = document.all.tags("BLINK")
	for (var i=0; i < blink.length; i++)
		blink[i].style.visibility = blink[i].style.visibility == "" ? "hidden" : ""
    }

function startBlink() 
{
if (document.all)
	setInterval("doBlink()",1000)
}

window.onload = startBlink;



function valFrmConf()
{
	if(document.confirmar.ResponseDiscount.value == '')
	{
		alert('<%=getofertAgentContraOfertLngStr("LtxtValRespDat")%>'.replace('{0}', '<%=LCase(txtOfert)%>'));
		return false;
	}
	else if (parseFloat(getNumericVB(document.confirmar.ResponseDiscount.value.replace('%',''))) <= 0)
	{
		alert('<%=getofertAgentContraOfertLngStr("LtxtValRespDat")%>'.replace('{0}', '<%=LCase(txtOfert)%>'));
		return false;
	}
	<% If Request("status") = "O" Then %>
	else if (parseInt(getNumericVB(document.confirmar.ResponseDiscount.value.replace('%',''))) == <%=rs("responseDiscount")%> && parseInt(getNumericVB(document.confirmar.ResponseQuantity.value)) == <%=rs("responseQuantity")%>)
	{
		alert('<%=getofertAgentContraOfertLngStr("LtxtNoChange")%>'.replace('{0}', '<%=LCase(txtOfert)%>'));
		return false;
	}
	<% end if %>
	
	return true;
}
</script>
<form method="POST" action="cart/ofertSubmit.asp" name="confirmar" onsubmit="return valFrmConf();">
	<div align="left">
		<table border="0" cellpadding="0" width="680">
			<tr>
				<td valign="top">
				<table border="0" cellpadding="0" width="100%">
					<% If tblCustTtl = "" Then %>
					<tr>
						<td id="tdMyTtl" class="FirmTlt" colspan="2"><%=rs("ItemName")%> </td>
					</tr>
					<% Else %>
					<% AddPath = "" %>
					<%=Replace(Replace(tblCustTtl, "{txtTitle}", rs("ItemName")), "{AddPath}", "")%>
					<% End If %>
					<tr class="FirmTlt3">
						<td style="height: 100px; vertical-align: top; padding-top: 2px;">
						<% If rs("UserText") <> "" Then %><%=rs("UserText")%><% End If %>&nbsp;
						</td>
						<td valign="top" style="padding-top: 2px; width: 100px;">
						<p align="center"><% If Pic <> "n_a.gif" then %><a href="javascript:Start('thumb/?item=<%=CleanItem(rs("ItemCode"))%>&pop=Y&AddPath=../',529,510,'yes')"><% end if %>
						<img border="1" src='pic.aspx?filename=<%=Pic%>&amp;dbName=<%=Session("olkdb")%>' class="ImgBorder"><% If Pic <> "n_a.gif" then %></a><% end if %></p>
						</td>
					</tr>
				</table>
				<table border="0" cellpadding="0" width="100%">
					<tr>
						<td>
						<table border="0" cellpadding="0" width="100%">
							<tr class="GeneralTbl">
								<td class="GeneralTblBold2" width="113">
								<%=getofertAgentContraOfertLngStr("DtxtCode")%></td>
								<td><%=rs("ItemCode")%>&nbsp;</td>
							</tr>
							<tr class="GeneralTbl">
								<td class="GeneralTblBold2"  width="113">
								<%=getofertAgentContraOfertLngStr("DtxtSalUnit")%>:</td>
								<td><%=SaleUn%>&nbsp;</td>
							</tr>
							<tr class="GeneralTbl">
								<td class="GeneralTblBold2"  width="113">
								<%=getofertAgentContraOfertLngStr("LtxtBasePrice")%><% If UnPrice <> "" Then %> (<%=UnPrice%>)<% End If %>:</td>
								<td><nobr><%=myApp.MainCur%>&nbsp;<%=FormatNumber(BasePrice,myApp.PriceDec)%></nobr></td>
							</tr>
							<tr class="GeneralTbl">
								<td class="GeneralTblBold2"  width="113">
								<%=getofertAgentContraOfertLngStr("DtxtState")%></td>
								<td><% Select Case Request("Status") 
                          Case "W" %>
                          <blink><font color="#FF9933"><%=getofertAgentContraOfertLngStr("DtxtWaiting")%></font></blink>
                   <%     Case "A" %>
                          <blink><font color="#008080"><%=getofertAgentContraOfertLngStr("DtxtAproved")%></font></blink>
                   <%     Case "O" %>
                          <blink><font color="#3366CC"><%=getofertAgentContraOfertLngStr("DtxtCounter")%>&nbsp;<%=txtOfert%></font></blink>
                   <%     Case "R" %>
                          <blink><font color="#FF0066"><%=getofertAgentContraOfertLngStr("DtxtRejected")%></font></blink>
                   <%     Case "C" %>
                          <blink><font color="#666699"><%=getofertAgentContraOfertLngStr("DtxtAnuled")%></font></blink>
                   <%     End Select %></td>
							</tr>
						</table>
						</td>
					</tr>
					<tr>
						<td>
						<table border="0" cellpadding="0" width="100%">
							<tr class="GeneralTbl">
								<td class="GeneralTblBold2" align="center" style="width: 50%">
								<%=Replace(Replace(getofertAgentContraOfertLngStr("LttlLastOffer"), "{0}", getofertAgentContraOfertLngStr("LtxtLast")), "{1}", txtOfert) %>
								</td>
								<td align="center" style="width: 50%"><%=getofertAgentContraOfertLngStr("LtxtResTo")%>&nbsp;<%=txtOfert%></td>
							</tr>
							<tr class="GeneralTbl">
								<td class="GeneralTblBold2" style="width: 50%">
								<table border="0" cellpadding="0" width="100%" cellspacing="1">
									<tr class="GeneralTblBold2">
										<td><%=getofertAgentContraOfertLngStr("LtxtLimDat")%>:</td>
										<td align="right" style="width: 70px"><%=FormatDate(rs("ofertLimit"), True)%></td>
									</tr>
								</table>
								</td>
								<td style="width: 50%">
								<table border="0" cellpadding="0" width="100%" cellspacing="1">
									<tr class="GeneralTbl">
										<td width="83"><%=getofertAgentContraOfertLngStr("LtxtLimit")%>:</td>
										<td align="right">
										<input <% If Request("status") = "C" Then %>disabled<% end if %> name="ResponseLimit" size="17" style="float: left; text-align:right" value="<% If Request("status") = "O" then %><%=rs("responseLimit")%><%end if%>" onfocus="this.select()" onchange="javascript:chkNum(this,'')"></td>
									</tr>
								</table>
								</td>
							</tr>
							<tr class="GeneralTbl">
								<td class="GeneralTblBold2" style="width: 50%">
								<table border="0" cellpadding="0" width="100%" cellspacing="1">
									<tr class="GeneralTblBold2">
										<td><%=getofertAgentContraOfertLngStr("DtxtPrice")%>:</td>
										<td align="right" style="width: 70px"><% If CDbl(rs("ofertPrice")) <> 0 then %><nobr><%=myApp.MainCur%>&nbsp;<%=FormatNumber(ofertPrice,myApp.PriceDec)%></nobr><% Else %>&nbsp;<% end if %></td>
									</tr>
								</table>
								</td>
								<td style="width: 50%">
								<table border="0" cellpadding="0" width="100%" cellspacing="1">
									<tr class="GeneralTbl">
										<td width="83"><%=getofertAgentContraOfertLngStr("DtxtPrice")%>:</td>
										<td align="right">
										<input <% If Request("status") = "C" Then %>disabled<% end if %> name="ResponsePrice" size="17" style="float: left; text-align:right" value="<% If Request("status") = "O" then %><%=FormatNumber(responsePrice,myApp.PriceDec)%><%end if%>" onfocus="this.select()" onchange="javascript:chkThis(this); setTotal()"></td>
									</tr>
								</table>
								</td>
							</tr>
							<tr class="GeneralTbl">
								<td class="GeneralTblBold2" style="width: 50%">
								<table border="0" cellpadding="0" width="100%" cellspacing="1">
									<tr class="GeneralTblBold2">
										<td><%=getofertAgentContraOfertLngStr("DtxtDiscount")%>:</td>
										<td style="width: 70px">
										<p align="right"><% If CDbl(rs("ofertDiscount")) <> "" Then %>%&nbsp;<%=FormatNumber(rs("ofertDiscount"),myApp.PercentDec)%><% Else %>&nbsp;<% end if %></td>
									</tr>
								</table>
								</td>
								<td style="width: 50%">
								<table border="0" cellpadding="0" width="100%" cellspacing="1">
									<tr class="GeneralTbl">
										<td width="83"><%=getofertAgentContraOfertLngStr("DtxtDiscount")%>:</td>
										<td align="right">
										<input <% If Request("status") = "C" Then %>disabled<% end if %> name="ResponseDiscount" size="17" style="float: left; text-align:right" value="<% If Request("status") = "O" then %><%=FormatNumber(rs("responseDiscount"),myApp.PercentDec)%><%end if%>" onfocus="this.select()" onchange="javascript:FormatQty(this,<%=myApp.PercentDec%>);changeDiscount(this); setTotal()"></td>
									</tr>
								</table>
								</td>
							</tr>
							<tr class="GeneralTbl">
								<td class="GeneralTblBold2" style="width: 50%">
								<table border="0" cellpadding="0" width="100%" cellspacing="1">
									<tr class="GeneralTblBold2">
										<td><%=getofertAgentContraOfertLngStr("DtxtQty")%>:</td>
										<td align="right" style="width: 70px">
										<% If CDbl(rs("ofertQuantity")) <> 0 then %><%=FormatNumber(ofertQuantity,myApp.QtyDec)%><% Else %>&nbsp;<% end if %></td>
									</tr>
								</table>
								</td>
								<td style="width: 50%">
								<table border="0" cellpadding="0" width="100%" cellspacing="1">
									<tr class="GeneralTbl">
										<td width="83"><%=getofertAgentContraOfertLngStr("DtxtQty")%>:</td>
										<td align="right">
										<input <% If Request("status") = "C" Then %>disabled<% end if %> name="ResponseQuantity" size="17" style="float: left; text-align:right" value="<% If Request("status") = "O" then %><%=FormatNumber(responseQuantity,myApp.QtyDec)%><%end if%>" onchange="javascript:FormatQty(this,<%=myApp.QtyDec%>);chkNum(this,1); setTotal()" onfocus="this.select()"></td>
									</tr>
								</table>
								</td>
							</tr>
							<tr class="GeneralTbl">
								<td class="GeneralTblBold2" style="width: 50%">
								<table border="0" cellpadding="0" width="100%" cellspacing="1">
									<tr class="GeneralTblBold2">
										<td><%=getofertAgentContraOfertLngStr("DtxtTotal")%>:</td>
										<td align="right" style="width: 70px">
										<% If CDbl(rs("ofertQuantity")) <> 0 then %><nobr><%=myApp.MainCur%>&nbsp;<%=FormatNumber(ofertTotal,myApp.PriceDec)%></nobr><% Else %>&nbsp;<% end if %></td>
									</tr>
								</table>
								</td>
								<td style="width: 50%">
								<table border="0" cellpadding="0" width="100%" cellspacing="1">
									<tr class="GeneralTbl">
										<td width="83"><%=getofertAgentContraOfertLngStr("DtxtTotal")%>:</td>
										<td align="right">
										<input <% If Request("status") = "C" Then %>disabled<% end if %> name="total" size="17" style="float: left; text-align:right" value="<% If Request("status") = "O" then %><%=FormatNumber(responseTotal,myApp.PriceDec)%><%end if%>" onchange="javascript:changeTotal(this)" onfocus="this.select()"></td>
									</tr>
								</table>
								</td>
							</tr>
						</table>
						</td>
					</tr>
					<tr class="GeneralTblBold2">
						<td><%=getofertAgentContraOfertLngStr("DtxtNote")%>:</td>
					</tr>
					<tr class="GeneralTblBold2">
						<td>
						<table border="0" cellpadding="0" width="100%">
							<tr>
								<td class="GeneralTblBold2"><%=txtOfert%></td>
								<td class="GeneralTbl"><% If Not IsNull(rs("OfertNote")) Then %><%=rs("OfertNote")%><% End If %>&nbsp;</td>
							</tr>
							<tr>
								<td class="GeneralTblBold2"><%=getofertAgentContraOfertLngStr("LtxtResp")%> </td>
								<td class="GeneralTbl">
								<input <% If Request("status") = "C" Then %>disabled<% end if %> type="text" name="ResponseNote" size="80" value="<% If Request("status") = "O" then %><%=myHTMLEncode(rs("ResponseNote"))%><%end if%>"></td>
							</tr>
						</table>
						</td>
					</tr>
					<tr>
						<td>
						<table border="0" cellpadding="0" width="100%">
							<tr>
							<% If Request("status") <> "C" Then %>
			                <% If Request("status") <> "A" then %>
								<td>
								<p align="center">
								<input type="button" value=" <%=getofertAgentContraOfertLngStr("DtxtAccept")%>&nbsp;<%=txtOfert%> " name="btnAccept" style="width: 110" onclick="javascript:<% If Request("status") = "R" Then %>if(confirm('<%=getofertAgentContraOfertLngStr("LtxtConfAccRefOfr")%>'.replace('{0}', '<%=LCase(txtOfert)%>')))<% end if %>window.location.href='cart/ofertSubmit.asp?cmd=acceptOfert&ofertIndex=<%=Request("ofertIndex")%><%=fltParam%>&page=<%=Request("page")%>&redir=<%=Request("redir")%>'"></td>
							<% End If %>
							<% If Request("status") <> "R" then %>
								<td>
								<p align="center">
								<input type="button" value="<%=getofertAgentContraOfertLngStr("DtxtReject")%>&nbsp;<%=txtOfert%>" name="btnReject" style="width: 110" onclick="javascript:if(confirm('<%=getofertAgentContraOfertLngStr("LtxtConfRefOfert")%>'.replace('{0}', '<%=LCase(txtOfert)%>')))window.location.href='cart/ofertSubmit.asp?cmd=rejectOfert&ofertIndex=<%=Request("ofertIndex")%><%=fltParam%>&page=<%=Request("page")%>&redir=<%=Request("redir")%>'"></td>
							<% End If %>
								<td>
								<p align="center">
								<% 
								If Request("status") = "O" then
									btnCounterOffer = getofertAgentContraOfertLngStr("DtxtUpdate")
								Else
									btnCounterOffer = getofertAgentContraOfertLngStr("DtxtCounter")
								End If
								%>
								<input type="submit" value="<%=Replace(Replace(getofertAgentContraOfertLngStr("LbtnCounterOffer"), "{0}", btnCounterOffer), "{1}", txtOfert)%>" name="I2" style="width: 110"></td>
							<% end if %>
								<td>
								<p align="center">
								<input type="button" value="<%=getofertAgentContraOfertLngStr("LtxtViewHist")%>" name="btnViewHist" style="width: 110" onclick="javascript:window.location.href='ofertHistory.asp?ofertIndex=<%=Request("ofertIndex")%>'"></td>
							</tr>
						</table>
						</td>
					</tr>
				</table>
				</td>
			</tr>
		</table>
	</div>
<input type="hidden" name="cmd" value="AgentContraOfert">
<input type="hidden" name="ofertIndex" value="<%=Request("ofertIndex")%>">
<% For each item in Request("flt") %>
<input type="hidden" name="flt" value="<%=item%>">
<% next %>
<input type="hidden" name="page" value="<%=Request("page")%>">
<input type="hidden" name="redir" value="<%=Request("redir")%>">
</form>
<% set rs = nothing %>
<!--#include file="agentBottom.asp"-->
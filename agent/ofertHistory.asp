<!--#include file="clientInc.asp"-->
<!--#include file="expire.inc"-->
<% Select Case userType
Case "C" %><!--#include file="clientTop.asp"-->
<% 
If Session("UserName") = "-Anon-" or not optOfert Then Response.Redirect "default.asp"
Case "V" %><!--#include file="agentTop.asp"-->
<% 
End Select
addLngPathStr = "" %>
<!--#include file="lang/ofertHistory.asp" -->
<script language="javascript">
function Start(page, w, h, s) {
OpenWinThumb = this.open(page, "myOlkThumb", "toolbar=no,menubar=no,location=no,scrollbars="+s+",resizable=yes, width="+w+",height="+h);
}
</script>
<%
Dim errmsg
set rs = Server.CreateObject("ADODB.recordset")
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetOfferHist" & Session("ID")
cmd.Parameters.Refresh()
cmd("@OfertIndex") = Request("ofertIndex")
If Request("linenum") <> "" Then cmd("@LineNum") = Request("linenum")
cmd("@LanID") = Session("LanID")
set rs = cmd.execute()

linenum = rs("LineNum")
SaleType = rs("SaleType")
If rs("PicturName") <> "" Then Pic = rs("PicturName") Else Pic = "n_a.gif"

First = 0
Last = CInt(rs("cant"))-1
If linenum = 0 then oPrev = Last Else oPrev = linenum-1
If CStr(linenum) = CStr(Last) then oNext = 0 Else oNext = linenum+1

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
<script language="javascript">
function changeScreenSize(w,h) { window.resizeTo( w,h ) }
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
</script>
<form method="POST" action="cart/addcartsubmitm.asp">
<table border="0" cellpadding="0" width="100%">
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
					<% If rs("UserText") <> "" Then %><%=rs("UserText")%><% End If %>
					</td>
					<td valign="top" style="padding-top: 2px; width: 100px;">
					<p align="center"><% If Pic <> "n_a.gif" then %><a href="javascript:Start('thumb/?item=<%=CleanItem(rs("ItemCode"))%>&pop=Y&AddPath=../',529,510,'yes')"><% end if %>
					<img border="1" src='pic.aspx?filename=<%=Pic%>&amp;dbName=<%=Session("olkdb")%>' class="ImgBorder"><% If Pic <> "n_a.gif" then %></a><% end if %></p>
					</td>
				</tr>
			</table>
			<table border="0" cellpadding="0" width="100%">
				<tr class="FirmTbl">
					<td class="FirmTlt3" width="39%"><%=getofertHistoryLngStr("DtxtState")%>: <blink><% Select Case rs("ofertStatus") 
                          Case "W" %>
                          <font color="#FF9933"><%=getofertHistoryLngStr("DtxtWaiting")%></font>
                    <%    Case "A" %> 
                          <font color="#008080"><%=getofertHistoryLngStr("DtxtAproved")%></font>
                    <%    Case "O" %>
                          <font color="#3366CC"><blink><%=Replace(Replace(getofertHistoryLngStr("LformatCounterOffer"), "{0}", getofertHistoryLngStr("DtxtCounter")), "{1}", Server.HTMLEncode(txtOfert)) %></blink></font>
                    <%    Case "R" %>
                          <font color="#FF0066"><%=getofertHistoryLngStr("DtxtReject")%></font>
                    <%    Case "C" %>
                          <font color="#666699"><%=getofertHistoryLngStr("DtxtAnuled")%></font>
                    <%    End Select %></blink></td>
					<td width="60%"><hr size="1"></td>
				</tr>
				<tr>
					<td colspan="2">
					<table border="0" cellpadding="0" width="100%" cellspacing="1">
						<tr>
							<td class="FirmTlt3" width="50%"><%=getofertHistoryLngStr("DtxtCode")%>:</td>
							<td class="FirmTbl" width="50%"><%=rs("ItemCode")%>&nbsp;</td>
						</tr>
						<tr>
							<td class="FirmTlt3"width="51%"><%=getofertHistoryLngStr("DtxtSalUnit")%>:</td>
							<td class="FirmTbl" width="50%"><%=SaleUn%>&nbsp;</td>
						</tr>
						<tr>
							<td class="FirmTlt3" width="50%"><%=getofertHistoryLngStr("LtxtBasePrice")%><% If UnPrice <> "" Then %> (<%=UnPrice%>)<% End If %></td>
							<td class="FirmTbl" width="50%"><nobr><%=myApp.MainCur%>&nbsp;<%=FormatNumber(BasePrice,myApp.PriceDec)%></nobr></td>
						</tr>
						<tr>
							<td class="FirmTlt3" width="50%">
							<p align="center"><%=Server.HTMLEncode(txtOfert)%></td>
							<td class="FirmTlt3" width="50%">
							<p align="center"><%=getofertHistoryLngStr("LtxtResp")%></td>
						</tr>
						<tr>
							<td class="FirmTlt3" width="50%">
							<table border="0" cellpadding="0" width="100%" cellspacing="1">
								<tr>
									<td class="FirmTlt3" width="82">
									<%=getofertHistoryLngStr("DtxtDate")%>:</td>
									<td class="FirmTbl">
									<p align="right"><%=FormatDate(rs("ofrDate"), True)%></td>
								</tr>
							</table>
							</td>
							<td class="FirmTbl" width="50%" >
							<table border="0" cellpadding="0" width="100%" cellspacing="1">
								<tr>
									<td class="FirmTlt3" width="77">
									<%=getofertHistoryLngStr("DtxtDate")%>:</td>
									<td class="FirmTbl">
									<p align="right">
									&nbsp;<%=FormatDate(rs("resDate"), True)%></td>
								</tr>
							</table>
							</td>
						</tr>
						<tr>
							<td class="FirmTlt3" width="50%">
							<table border="0" cellpadding="0" width="100%" cellspacing="1">
								<tr>
									<td class="FirmTlt3" width="82">
									<%=getofertHistoryLngStr("LtxtLimit")%>:</td>
									<td class="FirmTbl">
									<p align="right">&nbsp;<%=FormatDate(rs("ofertLimit"), True)%></td>
								</tr>
							</table>
							</td>
							<td class="FirmTbl" width="50%">
							<table border="0" cellpadding="0" width="100%" cellspacing="1">
								<tr>
									<td class="FirmTlt3" width="77">
									<%=getofertHistoryLngStr("LtxtLimit")%>:</td>
									<td class="FirmTbl">
									<p align="right">
									&nbsp;<%=FormatDate(rs("responseLimit"), True)%></td>
								</tr>
							</table>
							</td>
						</tr>
						<tr>
							<td class="FirmTlt3"width="51%">
							<table border="0" cellpadding="0" width="100%" cellspacing="1">
								<tr>
									<td class="FirmTlt3" width="82">
									<%=getofertHistoryLngStr("LtxtHour")%>:</td>
									<td class="FirmTbl">
									<p align="right"><%=rs("ofrHour")%></td>
								</tr>
							</table>
							</td>
							<td class="FirmTbl" width="50%">
							<table border="0" cellpadding="0" width="100%" cellspacing="1">
								<tr>
									<td class="FirmTlt3" width="77">
									<%=getofertHistoryLngStr("LtxtHour")%>:</td>
									<td class="FirmTbl">
									<p align="right">
									&nbsp;<%=rs("resHour")%></td>
								</tr>
							</table>
							</td>
						</tr>
						<tr>
							<td class="FirmTlt3" width="50%">
							<table border="0" cellpadding="0" width="100%" cellspacing="1">
								<tr>
									<td class="FirmTlt3" width="82">
									<%=getofertHistoryLngStr("DtxtPrice")%>:</td>
									<td class="FirmTbl">
									<p align="right"><nobr><%=myApp.MainCur%>&nbsp;<%=FormatNumber(ofertPrice,myApp.PriceDec)%></nobr></td>
								</tr>
							</table>
							</td>
							<td class="FirmTbl" width="50%">
							<table border="0" cellpadding="0" width="100%" cellspacing="1">
								<tr>
									<td class="FirmTlt3" width="77">
									<%=getofertHistoryLngStr("DtxtPrice")%>:</td>
									<td class="FirmTbl">
									<p align="right">
									&nbsp;<% If CDbl(rs("responsePrice")) <> 0 then %><nobr><%=myApp.MainCur%>&nbsp;<%=FormatNumber(responsePrice,myApp.PriceDec)%></nobr><% end if %></td>
								</tr>
							</table>
							</td>
						</tr>
						<tr>
							<td class="FirmTlt3" width="50%">
							<table border="0" cellpadding="0" width="100%" cellspacing="1">
								<tr>
									<td class="FirmTlt3" width="82">
									<%=getofertHistoryLngStr("DtxtDiscount")%>:</td>
									<td class="FirmTbl">
									<p align="right">%&nbsp;<%=FormatNumber(rs("ofertDiscount"),myApp.PercentDec)%></td>
								</tr>
							</table>
							</td>
							<td class="FirmTbl" width="50%">
							<table border="0" cellpadding="0" width="100%" cellspacing="1">
								<tr>
									<td class="FirmTlt3" width="77">
									<%=getofertHistoryLngStr("DtxtDiscount")%>:</td>
									<td class="FirmTbl">
									<p align="right">&nbsp;<% if CDbl(rs("responseDiscount")) <> 0 then %>%&nbsp;<%=FormatNumber(rs("responseDiscount"),myApp.PercentDec)%><% end if %></td>
								</tr>
							</table>
							</td>
						</tr>
						<tr>
							<td class="FirmTlt3" width="50%">
							<table border="0" cellpadding="0" width="100%" cellspacing="1">
								<tr>
									<td class="FirmTlt3" width="82">
									<%=getofertHistoryLngStr("DtxtQty")%>:</td>
									<td class="FirmTbl">
									<p align="right"><%=FormatNumber(ofertQuantity,myApp.QtyDec)%></td>
								</tr>
							</table>
							</td>
							<td class="FirmTbl" width="50%">
							<table border="0" cellpadding="0" width="100%" cellspacing="1">
								<tr>
									<td class="FirmTlt3" width="76">
									<%=getofertHistoryLngStr("DtxtQty")%>:</td>
									<td class="FirmTbl">
									<p align="right">&nbsp;<% If CDbl(rs("responseQuantity")) <> 0 then %><%=FormatNumber(responseQuantity,myApp.QtyDec)%><% end if %></td>
								</tr>
							</table>
							</td>
						</tr>
						<tr>
							<td class="FirmTlt3" width="50%">
							<table border="0" cellpadding="0" width="100%" cellspacing="1">
								<tr>
									<td class="FirmTlt3" width="82">
									<%=getofertHistoryLngStr("DtxtTotal")%>:</td>
									<td class="FirmTbl">
									<p align="right">&nbsp;<nobr><%=myApp.MainCur%>&nbsp;<%=FormatNumber(ofertTotal,myApp.PriceDec)%></nobr></td>
								</tr>
							</table>
							</td>
							<td class="FirmTbl" width="50%">
							<table border="0" cellpadding="0" width="100%" cellspacing="1">
								<tr>
									<td class="FirmTlt3" width="75">
									<%=getofertHistoryLngStr("DtxtTotal")%>:</td>
									<td class="FirmTbl">
									<p align="right"><% If CDbl(rs("responseQuantity")) <> 0 then %><nobr><%=myApp.MainCur%>&nbsp;<%=FormatNumber(responseTotal,myApp.PriceDec)%></nobr><% end if %></td>
								</tr>
							</table>
							</td>
						</tr>
						<tr class="FirmTbl">
							<td colspan="2"><b><%=getofertHistoryLngStr("DtxtNote")%>&nbsp;<%=txtOfert%>:</b> <% If Not IsNull(rs("OfertNote")) Then %><%=rs("OfertNote")%><% End If %></td>
						</tr>
						<tr class="FirmTbl">
							<td colspan="2"><b><%=getofertHistoryLngStr("DtxtNote")%>&nbsp;<%=getofertHistoryLngStr("LtxtResp")%>:</b> <% If Not IsNull(rs("ResponseNote")) Then %><%=rs("ResponseNote")%><% End If %></td>
						</tr>
						<tr class="FirmTbl">
							<td colspan="2">
							&nbsp;</td>
						</tr>
						<tr class="FirmTbl">
							<td colspan="2">
							<p align="center">
					<input type="button" value="&lt;&lt;" name="B9"  onclick="javascript:window.location.href='ofertHistory.asp?ofertIndex=<%=Request("ofertIndex")%>&linenum=<%=First%>'"> 
					<input type="button" value="&lt;" name="B10" onclick="javascript:window.location.href='ofertHistory.asp?ofertIndex=<%=Request("ofertIndex")%>&linenum=<%=oPrev%>'"> 
					-
							<input type="button" value="&gt;" name="B11" onclick="javascript:window.location.href='ofertHistory.asp?ofertIndex=<%=Request("ofertIndex")%>&linenum=<%=oNext%>'"> 
					<input type="button" value="&gt;&gt;" name="B12" onclick="javascript:window.location.href='ofertHistory.asp?ofertIndex=<%=Request("ofertIndex")%>&linenum=<%=Last%>'"></td>
						</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td colspan="2">
					<table border="0" cellpadding="0" width="100%">
						<tr>
							<td style="width: 100px;">
							<p align="center">
		 					<input type="button" value="<%=getofertHistoryLngStr("DtxtBack")%>" class="btnBack" name="I4" style="width:100px;" onclick="javascript:history.go(-1);"></td>
		 					<td>&nbsp;</td>
			                <% If userType = "C" and ((rs("ofertstatus") = "A" and (IsNull(rs("ofertDays")) or rs("ofertDays") >= 0)) or (rs("ofertStatus") = "O" and (IsNull(rs("responseDays")) or rs("responseDays") >= 0))) Then %>
							<td style="width: 100%;">
							<p align="center">
			 				<input type="submit" class="btnBuy" value="<%=getofertHistoryLngStr("LtxtPurchase")%>" name="I2" style="width:100px;"></td>
								<td style="width: 100%;">
								<p align="center">
			  					<input class="btnNew" type="button" value="<%=Replace(Replace(getofertHistoryLngStr("LformatNewOffer"), "{0}", getofertHistoryLngStr("DtxtNew")), "{1}", Server.HTMLEncode(txtOfert))%>" name="I3" style="width:100" onclick="javascript:window.location.href='ofertContraOfert.asp?ofertIndex=<%=Request("ofertIndex")%>&pop=Y&AddPath=../'"></td>
			                <% end if %>
						</tr>
					</table>
					</td>
				</tr>
				</table>
			</td>
		</tr>
	</table>
<INPUT TYPE="HIDDEN" NAME="Item" VALUE="<%=CleanItem(rs("ItemCode"))%>">
<input type="hidden" name="precio" value="<% If rs("ofertStatus") = "A" then %><%=ofertPrice%><% ElseIf rs("ofertStatus") = "O" then %><%=responsePrice%><% end if %>">
<input type="hidden" name="T1" value="<% If rs("ofertStatus") = "A" then %><%=ofertQuantity%><% ElseIf rs("ofertStatus") = "O" then %><%=responseQuantity%><% end if %>">
<input type="hidden" name="redir" value="oferts">
<input type="hidden" name="SaleType" value="<%=SaleType%>">
<input type="hidden" name="OfertIndex" value="<%=Request("OfertIndex")%>">
<input type="hidden" name="AddPath" value="../">
<input type="hidden" name="pop" value="Y">
</form>

<% If setCustTtl Then %>
<script language="javascript" src="setTltBg.js.asp?custTtlBgL=<%=custTtlBgL%>&amp;custTtlBgM=<%=custTtlBgM%>&amp;AddPath=../"></script>
<script language="javascript">setTtlBg(false);</script>
<% End If %>
<% set rs = nothing %>
<% Select Case userType
Case "C" %><!--#include file="clientBottom.asp"-->
<% Case "V" %><!--#include file="agentBottom.asp"-->
<% End Select %>
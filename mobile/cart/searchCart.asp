<% addLngPathStr = "cart/" %>
<!--#include file="lang/searchCart.asp" -->
<div align="center">
	<center>
	<table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
		<tr>
			<td bgcolor="#9BC4FF">
			<table border="0" cellpadding="0" bordercolor="#111111" width="100%" id="AutoNumber1">
				<tr>
					<td width="100%">
					<p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><%=getsearchCartLngStr("LtxtSearchItmInCart")%></font></b></p>
					</td>
				</tr>
				<tr>
					<td width="100%">
					<form method="POST" action="operaciones.asp" name="search1">
						<table border="0" cellpadding="0" cellspacing="1" bordercolor="#111111" width="100%" id="AutoNumber2">
							<tr>
								<td width="100%" bgcolor="#66A4FF">
								<p align="center"><b>
								<font size="1" face="Verdana"><%=getsearchCartLngStr("LtxtSearchByItm")%></font></b></p>
								</td>
							</tr>
							<tr>
								<td width="100%">
								<p align="center">
								<font face="Verdana" size="2" color="#336699">
								<b>
								<input type="text" name="string" size="25" style="font-size:12px;"></b></font></p>
								</td>
							</tr>
							<tr>
								<td width="100%">
								<p align="left">
								<font face="Verdana" size="2" color="#336699">
								<b></b></font></p>
								</td>
							</tr>
							<% If 1 = 2 Then %>
							<tr>
								<td width="100%">
								<p align="center"><font face="Verdana">
								<input type="radio" value="E" name="rdSearchAs" id="rdSearchAsE" checked><label for="rdSearchAsE"><font size="2"><%=getsearchCartLngStr("LtxtExact")%></font></label><input type="radio" name="rdSearchAs" id="rdSearchAsS" value="S"><label for="rdSearchAsS"><font size="2"><%=getsearchCartLngStr("LtxtLike")%></font></label></font></p>
								</td>
							</tr>
							<% End If %>
							<tr>
								<td width="100%">
								<p align="center">
								<input type="submit" value="<%=getsearchCartLngStr("DbtnSearch")%>" name="B1"></p>
								</td>
							</tr>
						</table>
						<input type="hidden" name="cmd" value="cart">
						<input type="hidden" name="document" value="B">
					</form>
&nbsp;</td>
				</tr>
			</table>
			</td>
		</tr>
	</table>
	</center></div>
<% addLngPathStr = "ventas/" %><!--#include file="lang/maps.asp" -->
<% sql = "select AdresType, Address " & _
"from CRD1 T0 " & _
"where T0.CardCode = N'yuval' order by AdresType, Address "
rs.close
rs.open sql, conn, 3, 1 %>
<table border="0" cellpadding="0" style="width: 100%">
	<tr class="GeneralTlt">
		<td colspan="2"><%=getmapsLngStr("LtxtMaps")%></td>
	</tr>
	<tr>
		<td style="width: 200px; vertical-align: top;">
			<table border="0" cellpadding="0" width="100%">
				<tr>
					<td class="GeneralTblBold2"><%=getmapsLngStr("LtxtBillAddress")%></td>
				</tr>
				<% rs.Filter = "AdresType = 'B'"
				do while not rs.eof %>
				<tr class="GeneralTbl">
					<td><%=rs("Address")%></td>
				</tr>
				<% rs.movenext
				loop %>
				<tr>
					<td class="GeneralTblBold2"><%=getmapsLngStr("LtxtShipAddress")%></td>
				</tr>
				<% rs.Filter = "AdresType = 'S'"
				do while not rs.eof %>
				<tr class="GeneralTbl">
					<td><%=rs("Address")%></td>
				</tr>
				<% rs.movenext
				loop %>
			</table>
		</td>
		<% iFrameH = Session("sHeight")-700
		If iFrameH < 400 Then iFrameH = 400
		%>
		<td style="vertical-align: top;">
			<table style="width: 100%">
				<tr>
					<td style="height: 60px;" class="GeneralTbl">&nbsp;</td>
				</tr>
				<tr>
					<td class="GeneralTbl">
					<iframe width="100%" height="<%=iFrameH%>" frameborder="0" scrolling="no" marginheight="0" marginwidth="0" src="http://maps.google.com/maps?f=q&amp;hl=<%=Session("myLng")%>&amp;geocode=&amp;q=<%=CountryName%>&amp;ie=UTF8&amp;output=embed&amp;s=AARTsJocuAqeZw-TNAyU4Rj7YSqwSOos8Q"></iframe><br /><small><a href="http://maps.google.com/maps?f=q&hl=en&geocode=&q=<%=CountryName%>" class="LinkTop" target="_blank"><%=getmapsLngStr("LtxtViewLargeMap")%></a></small>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
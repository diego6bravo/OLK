<!--#include file="top.asp" -->
<!--#include file="lang/adminUpdate.asp" -->


<%

	set oLic = Server.CreateObject("TM.LicenceConnect.LicenceConnection")
	oLic.LicenceServer = licip
	oLic.LicencePort = "8633" 
	If oLic.IsAlive Then
		LicErr = False
		strHDKey = oLic.GetHDKey()
	Else
		LicErr = True
	End If


%>
<table border="0" cellpadding="0" cellspacing="2" width="100%" id="table1">
	<tr>
		<td>
		<p align="center">
		<img border="0" src="images/updates_img.jpg" width="453" height="164"></td>
	</tr>
	<tr>
		<td>
		<div align="center">
			<table border="0" cellpadding="0" width="550" id="table2">
				<tr>
					<td bgcolor="#F0F9FE"><b><font size="1" face="Verdana">
					<%=getadminUpdateLngStr("LttlUpdate")%></font></b></td>
				</tr>
				<% If Not LicErr Then %>
				<tr>
					<td>
					<iframe width=540 height=375 src="http://www.topmanage.com.pa/desarrollo/olk_update332632262625652626256/new_olk_download.asp?v=<%=OLKVerStr%>&HDKey=<%=myHTMLEncode(strHDKey)%>&myLng=<%=Session("myLng")%>" align=left frameborder=0 hspace=0 vspace=0 name="contenido" scrolling="no" style="border: 1px solid #000000"></iframe></td>
				</tr>
				<% Else %>
				<tr>
					<td bgcolor="#F0F9FE"><b><font size="1" face="Verdana">
					<%=getadminUpdateLngStr("LtxtErrTtl")%> 
					<br><%=getadminUpdateLngStr("LtxtErrTxt1")%>
					</font></b>
					<p><b><font size="1" face="Verdana">
					1 - <%=getadminUpdateLngStr("LtxtServStart")%><br>
					2 - <%=getadminUpdateLngStr("LtxtNoLic")%><br>
					3 - <%=getadminUpdateLngStr("LtxtLicExp")%> &nbsp;</font></b></td>
				</tr>
				<% End If %>
			</table>
		</div>
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
</table>
<!--#include file="bottom.asp" -->
<% addLngPathStr = "cart/" %>
<!--#include file="lang/itemVolRep.asp" -->
<%

sql = 	"select OITM.ItemCode, OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'ItemName', OITM.ItemCode, OITM.ItemName) ItemName, NumInSale, SalPackUn, " & _
		"Replace(Convert(nvarchar(4000),OLKCommon.dbo.DBOLKGetTrans" & Session("ID") & "(" & Session("LanID") & ", 'OITM', 'UserText', OITM.ItemCode, UserText)) collate database_default,Char(13),'<br>') Notes, PicturName " & _
		"from oitm " & _
		"where SellItem = 'Y' and oitm.ItemCode = N'" & Request("Item") & "'"
set rs = conn.execute(sql)

If rs("PicturName") <> "" Then
	Pic = rs("PicturName")
Else
	Pic = "n_a.gif"
End If

volSelBy = 1
If CInt(Request("un")) > 1 Then volSelBy = volSelBy * CDbl(rs("NumInSale"))
If CInt(Request("un")) = 3 Then volSelBy = volSelBy * CDbl(rs("SalPackUn"))


 %>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
    <tr>
      <td bgcolor="#9BC4FF">
        <table border="0" cellpadding="0" bordercolor="#111111" width="100%">
        <!--#include file="../C_Art/CardNameAdd.asp" -->
        <tr>
          <td>
          <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><%=DocName%>&nbsp;<%=getitemVolRepLngStr("DtxtItem")%>&nbsp;<%=Request("Item")%>
          </font></b></td>
        </tr>
        <tr>
          <td>
          <table border="0" cellpadding="0" cellspacing="1"  bordercolor="#111111" width="100%" id="AutoNumber2">
            <tr>
              <td width="86" valign="top">
                <a href="itemVolRep.asp?cmd=viewImage&amp;FileName=<%=Pic%>"><img border="0" src="pic.aspx?filename=<%=Pic%>&dbName=<%=Session("olkdb")%>&MaxSize=80"></a></td>
              <td valign="top" width="154"><font size="1" face="Verdana"><%=rs("Notes") %></font></td>
            </tr>
          </table>
          </td>
        </tr>
        <tr>
          <td>
          <table border="0" cellpadding="0"  bordercolor="#111111" width="100%">
            <tr>
              <td bgcolor="#75ACFF" align="left"><b><font size="1" face="Verdana"><%=getitemVolRepLngStr("DtxtCode")%></font></b></td>
              <td><font face="Verdana" size="1"><% =rs("ItemCode") %></font></td>
            </tr>
            <tr>
              <td bgcolor="#75ACFF" align="left" valign="top"><b><font size="1" face="Verdana"><%=getitemVolRepLngStr("DtxtDescription")%></font></b></td>
              <td><font face="Verdana" size="1"><% =rs("ItemName") %></font></td>
            </tr>
          </table>
          </td>
        </tr>
		<tr>
			<td>
				<table cellpadding="0" cellspacing="0" border="0" width="100%">
					<tr>
						<td valign="top">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<% 
							set rd = Server.CreateObject("ADODB.RecordSet")
							sql = "select T0.Amount,  " & _
								"Case When AutoUpdt = 'N' Then T0.Price  " & _
								"			Else T2.Price-(T2.Price*T0.Discount/100) " & _
								"	End Price " & _
								"from spp2 T0 " & _
								"inner join spp1 T1 on T1.ItemCode = T0.ItemCode and T1.CardCode = T0.CardCode and T1.LineNum = T0.SPP1LNum " & _
								"inner join ITM1 T2 on T2.ItemCode = T0.ItemCode and T2.PriceList = T1.ListNum " & _
								"where (T0.CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "' or T0.CardCode = N'*"  & Session("plist") &"' and not exists(select 'A' " & _
								"from spp2 T0 " & _
								"inner join spp1 T1 on T1.ItemCode = T0.ItemCode and T1.CardCode = T0.CardCode and T1.LineNum = T0.SPP1LNum " & _
								"inner join ITM1 T2 on T2.ItemCode = T0.ItemCode and T2.PriceList = T1.ListNum " & _
								"where (T0.CardCode = N'" & saveHTMLDecode(Session("UserName"), False) & "') and T0.ItemCode = N'" & saveHTMLDecode(Request("Item"), False) & "' " & _
								"and " & _
								"  (T1.FromDate is null or DateDiff(day,getdate(),T1.FromDate) <= 0) and " & _
								"  (T1.ToDate is null or DateDiff(day,getdate(),T1.ToDate) >= 0))) and T0.ItemCode = N'" & saveHTMLDecode(Request("Item"), False) & "' " & _
								"and " & _
								"  (T1.FromDate is null or DateDiff(day,getdate(),T1.FromDate) <= 0) and " & _
								"  (T1.ToDate is null or DateDiff(day,getdate(),T1.ToDate) >= 0) "
							set rd = conn.execute(sql) %>
							<tr>
								<td colspan="3">
								<table border="0" width="100%">
									<tr>
										<td colspan="2" bgcolor="#75ACFF" align="center"><b><font size="1" face="Verdana"><%=getitemVolRepLngStr("LtxtVolDiscount")%></font></b></td>
									</tr>
									<tr>
										<td width="50%" bgcolor="#75ACFF"><b><font size="1" face="Verdana"><%=getitemVolRepLngStr("DtxtQty")%></font></b></td>
										<td width="50%" bgcolor="#75ACFF"><b><font size="1" face="Verdana"><%=getitemVolRepLngStr("DtxtPrice")%></font></b></td>
									</tr>
									<% rd.movefirst
									do while not rd.eof %>
									<tr>
										<td width="50%" align="right" bgcolor="#75ACFF"><font size="1" face="Verdana"><%=rd("Amount")%></font></td>
										<td width="50%" align="right" bgcolor="#75ACFF"><font size="1" face="Verdana"><%=FormatNumber(CDbl(rd("Price"))*volSelBy, myApp.PriceDec)%></font></td>
									</tr>
									<% rd.movenext
									loop %>
								</table>
								</td>
							</tr>
							</table>	
						</td>
					</tr>
				</table>
			</td>
		</tr>
        </table>
      </td>
    </tr>
    </table>
  </center>
</div>
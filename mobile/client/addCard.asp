<% addLngPathStr = "client/" %><!--#include file="lang/addCard.asp" -->
<%
sql = "select CardCode, IsNull(CardName, '') CardName, CardType, IsNull(GroupCode,-1) GroupCode, IsNull(LicTradNum, '') LicTradNum, IsNull(CmpPrivate, '') CmpPrivate, IsNull(SlpCode, -1) SlpCode" & ", " & _
"(select Command from R3_ObsCommon..TLOG where LogNum = T0.LogNum) Command " & _
"from R3_ObsCommon..TCRD T0 " & _
"where T0.LogNum = " & Session("CrdRetVal")
set rs = conn.execute(sql)
Confirm = myAut.GetCardProperty(rs("CardType"), "C")
isUpdate = rs("Command") = "U"
If Request.Form.Count = 0 Then
	CardCode = rs("CardCode")
	CardType = rs("CardType")
	CardName = rs("CardName")
	CmpPrivate = rs("CmpPrivate")
	GroupCode = rs("GroupCode")
	SlpCode = rs("SlpCode")
	LicTradNum = rs("LicTradNum")
Else
	CardCode = Request("CardCode")
	CardType = Request("CardType")
	CardName = Request("CardName")
	CmpPrivate = Request("CmpPrivate")
	GroupCode = Request("GroupCode")
	SlpCode = Request("SlpCode")
	LicTradNum = Request("LicTradNum")
End If


Select Case CardType
	Case "S"
		GrpCardType = "S"
	Case "C", "L"
		GrpCardType = "C"
End Select %>
<script type="text/javascript">
<!--
function changeCardType(value)
{
	document.frmCrd.action='operaciones.asp';
	document.frmCrd.cmd.value = 'newClient';
	document.frmCrd.GroupCode.selectedIndex = -1;
	document.frmCrd.submit();

}
function save(redir)
{
	document.frmCrd.action='client/submitClient.asp';
	document.frmCrd.cmd.value = 'data';
	document.frmCrd.redir.value = redir;
	document.frmCrd.submit();
}

function changeCmpPrivate(value)
{
	<% If myApp.LawsSet = "MX" Then %>
	var maxLength = 12;
	if (value == 'I') maxLength = 13;
	
	document.frmCrd.LicTradNum.maxLength = maxLength;
	if (document.frmCrd.LicTradNum.value.length > maxLength)
		document.frmCrd.LicTradNum.value = document.frmCrd.LicTradNum.value.substring(0, maxLength);
	<% End If %>
}
//-->
</script>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
  <form method="post" action="client/submitClient.asp" name="frmCrd">
    <tr>
      <td>
      <img src="images/spacer.gif" width="100%" height="1" border="0" alt></td>
    </tr>
    <tr>
      <td bgcolor="#9BC4FF">
      <table border="0" cellpadding="0" bordercolor="#111111" width="100%" id="AutoNumber1">
        <tr>
          <td width="100%">
          <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1"><% If Not isUpdate Then %><%=getaddCardLngStr("LttlNewClient")%><% Else %><%=getaddCardLngStr("LttlEditClient")%><% End If %>
          </font></b></td>
        </tr>
        <tr>
          <td width="100%">
         <!--#include file="clientMenu.asp"--></td>
        </tr>
        <tr>
          <td width="100%">
            <table border="0" cellpadding="0"  bordercolor="#111111" width="100%" id="AutoNumber3">
            <tr>
              <td width="30%" bgcolor="#7DB1FF"><b>
              <font size="1" face="Verdana"><%=getaddCardLngStr("DtxtCode")%></font></b></td>
              <td width="70%" bgcolor="#8CBAFF">
              
              <table style="width: 100%" cellspacing="0" cellpadding="0">
							<tr>
											<td>
				<% If isUpdate Then %><input type="hidden" name="CardCode" value="<%=myHTMLEncode(CardCode)%>"><% End If %>
              <input type="text" <% If isUpdate or not isUpdate and myApp.AutoGenOCRD Then %>disabled<% End If %> name="CardCode<% If isUpdate Then %>2<% End If %>" size="15" value="<% If isUpdate or not isUpdate and not myApp.AutoGenOCRD Then %><%=myHTMLEncode(CardCode)%><% Else %><%=getaddCardLngStr("DtxtAutomatic")%><% End If %>" maxlength="15" <% If isUpdate Then %>class="InputDes"<% End If %>></td>
											<td><p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
				<select size="1" name="CardType" onchange="javascript:changeCardType(this.value);">
		      	<% If myAut.HasAuthorization(45) or isUpdate and CardType = "C" Then %><option <% If CardType = "C" Then %>selected<% End If %> value="C"><%=myHTMLEncode(txtClient)%></option><% End If %>
				<% If myAut.HasAuthorization(78) or isUpdate and CardType = "S" Then %><option <% If CardType = "S" Then %>selected<% End If %> value="S">
				<%=getaddCardLngStr("DtxtSupplier")%></option><% End If %>
				<% If myAut.HasAuthorization(77) or isUpdate and CardType = "L" Then %><option <% If CardType = "L" Then %>selected<% End If %> value="L">
				<%=getaddCardLngStr("DtxtLead")%></option><% End If %>
				</select></td>
							</tr>
				</table>
              </td>
            </tr>
		<tr>
              <td width="30%" bgcolor="#7DB1FF"><b>
              <font size="1" face="Verdana"><%=getaddCardLngStr("DtxtName")%></font></b></td>
              <td width="70%" bgcolor="#8CBAFF">
              <p align="center">
    <input type="text" name="CardName" size="40" value="<%=myHTMLEncode(CardName)%>" maxlength="100"></td>
            </tr>
            <% If InStr("MX, CR, GT, US, CA", myApp.LawsSet) <> 0 Then %>
			<tr>
              <td width="30%" bgcolor="#7DB1FF"><b>
              <font size="1" face="Verdana"><%=getaddCardLngStr("DtxtType")%></font></b></td>
              <td width="70%" bgcolor="#8CBAFF">
					<select size="1" name="CmpPrivate" onchange="changeCmpPrivate(this.value);">
				    <option value="C"><%=getaddCardLngStr("DtxtCmp")%></option>
					<option <% If CmpPrivate = "I" Then %>selected<% End If %> value="I">
					<%=getaddCardLngStr("LtxtNatPer")%></option>
				    </select></td>
            	</tr>
				<% Else %>
				<input type="hidden" name="CmpPrivate" value="C">
            	<% End If %>
				<tr>
              <td width="30%" bgcolor="#7DB1FF"><b>
              <font size="1" face="Verdana"><%=getaddCardLngStr("DtxtGroup")%></font></b></td>
              <td width="70%" bgcolor="#8CBAFF">
              <p>
    		<select size="1" name="GroupCode">
				<% 
				set rd = Server.CreateObject("ADODB.RecordSet")
				set cmd = Server.CreateObject("ADODB.Command")
				cmd.ActiveConnection = connCommon
				cmd.CommandType = &H0004
				cmd.CommandText = "DBOLKGetCrdGroups" & Session("ID")
				cmd.Parameters.Refresh()
				cmd("@LanID") = Session("LanID")
				If CardType = "S" Then cmd("@CardType") = "S" Else cmd("@CardType") = "C"
				set rd = cmd.execute()
				do While NOT rd.EOF %>
				<option <% If CStr(rd(0)) = CStr(GroupCode) Then %>selected<% End If %> value="<%=rd(0)%>"><%=myHTMLEncode(rd(1))%></option>
				<% rd.movenext
				loop %>
				</select></td>
            	</tr>
				<tr>
              <td width="30%" bgcolor="#7DB1FF"><b>
              <font size="1" face="Verdana"><% 
				Select Case myApp.LawsSet
					Case "PA", "IL", "US", "CA", "AT", "AU", "BE", "CH", "CZ", "DE", "DK", "ES", "FI", "FR", "CN", "CY", "HU", "IT", "NL", "NO", "PL", "PT", "RU", "SE", "SK", "GB", "ZA" %><%=getaddCardLngStr("DtxtLicTradNum")%><% 
					Case "MX", "CR", "GT" %>RFC<% 
					Case "CL" %>RUT<% 
					Case "BR" %>CRPJ<%
				End Select
				MaxLength = 32
				If myApp.LawsSet = "MX" Then
					Select Case rs("CmpPrivate")
						Case "I"
							MaxLength = 13
						Case Else
							MaxLength = 12
					End Select
				End If %></font></b></td>
              <td width="70%" bgcolor="#8CBAFF">
              <input type="text" name="LicTradNum" id="LicTradNum" size="40" value="<%=myHTMLEncode(LicTradNum)%>" maxlength="<%=MaxLength%>"></td>
            	</tr>
				<tr>
              <td width="30%" bgcolor="#7DB1FF"><b>
              <font size="1" face="Verdana"><%=getaddCardLngStr("DtxtAgent")%></font></b></td>
              <td width="70%" bgcolor="#8CBAFF">
              <select size="1" name="SlpCode">
				<%
				set cmd = Server.CreateObject("ADODB.Command")
				cmd.ActiveConnection = connCommon
				cmd.CommandType = &H0004
				cmd.CommandText = "DBOLKGetAgents" & Session("ID")
				cmd.Parameters.Refresh()
				cmd("@LanID") = Session("LanID")
				set rd = cmd.execute()
                do while not rd.eof %>
                <option value="<%=rd("SlpCode")%>" <% If rd("SlpCode") = -1 and IsNull(SlpCode) or CInt(SlpCode) = CInt(rd("SlpCode")) Then %>selected<% End If %>><%=myHTMLEncode(rd("SlpName"))%></option>
                <% rd.movenext
                loop %>
				</select></td>
            	</tr>
            <tr>
              <td bgcolor="#7DB1FF" colspan="2">
              <table cellpadding="0" border="0" width="100%">
				<tr>
					<% If not isUpdate Then %><td align="center"><input border="0" src="images/save_icon.gif" name="btnSave" type="image">
					</td><% End If %><% If isUpdate Then btnSubmit = "btnUpdate" Else btnSubmit = "btnAdd" %>
					<td align="center"><input type="image" name="<%=btnSubmit%>" value="<%=btnSubmit%>" border="0" src="images/ok_icon.gif"></td>
					<td align="center"><a href="operaciones.asp?cmd=clientcancel"><img border="0" src="images/x_icon.gif"></a>
					</td>
				</tr>
			</table>
			</td>
            </tr>
          </table>
           </td>
        </tr>
        <tr>
          <td width="100%">&nbsp;</td>
        </tr>
      </table>
      </td>
    </tr>
	<input type="hidden" name="cmd" value="data">
	<input type="hidden" name="confirm" value="">
	</form>
    </table>
  </center>
</div>
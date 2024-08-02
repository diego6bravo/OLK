<!--#include file="top.asp" -->
<!--#include file="lang/adminPriceCod.asp" -->

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style type="text/css">
.style1 {
	font-family: Verdana;
	font-size: xx-small;
	color: #3F7B96;
}
</style>
</head>

<% conn.execute("use [" & Session("OLKDB") & "]")
If Request("codType") <> "" Then codType = Request("codType") Else codType = "L"
sql = "select " & _
"(select NewKey from OLKMyCod where Type = '" & codType & "' and OrgKey = '0') '0', " & _
"(select NewKey from OLKMyCod where Type = '" & codType & "' and OrgKey = '1') '1', " & _
"(select NewKey from OLKMyCod where Type = '" & codType & "' and OrgKey = '2') '2', " & _
"(select NewKey from OLKMyCod where Type = '" & codType & "' and OrgKey = '3') '3', " & _
"(select NewKey from OLKMyCod where Type = '" & codType & "' and OrgKey = '4') '4', " & _
"(select NewKey from OLKMyCod where Type = '" & codType & "' and OrgKey = '5') '5', " & _
"(select NewKey from OLKMyCod where Type = '" & codType & "' and OrgKey = '6') '6', " & _
"(select NewKey from OLKMyCod where Type = '" & codType & "' and OrgKey = '7') '7', " & _
"(select NewKey from OLKMyCod where Type = '" & codType & "' and OrgKey = '8') '8', " & _
"(select NewKey from OLKMyCod where Type = '" & codType & "' and OrgKey = '9') '9', " & _
"(select NewKey from OLKMyCod where Type = '" & codType & "' and OrgKey = '.') '.', " & _
"(select NewKey from OLKMyCod where Type = '" & codType & "' and OrgKey = '-') '-' "
set rs = conn.execute(sql) %>
<script language="javascript">
<!--
function valFrm()
{
	 myFields = document.frmMyCod.FldCod;
	 for (var i = 0;i<myFields.length;i++)
	 {
	 	for (var j = 0;j<myFields.length;j++)
	 	{
	 		if (i != j) if (myFields[i].value == myFields[j].value) { alert('<%=getadminPriceCodLngStr("LtxtValLtr")%>'.replace('{0}', myFields[j].value)); return false; }
	 	}
	 }
	 return true;
}
//-->
</script>
<div align="center">
<img border="0" src="images/top_seguridad_2.jpg">
      <table border="0" cellpadding="0" bordercolor="#111111" width="163" id="table1" height="113" >
        <form method="POST" name="frmMyCod" action="adminSubmit.asp" onsubmit="return valFrm();">
        <tr>
          <td colspan="4" bgcolor="#E1F3FD" width="159" height="10"><b>
			<font face="Verdana" size="1" color="#31659C"><%=getadminPriceCodLngStr("LttlNumCod")%></font></b></td>
        </tr>
        <tr>
          <td colspan="4" bgcolor="#E8F5FD" width="159" height="9">
			<select size="1" name="codType" class="input" style="width: 100%" onchange="javascript:window.location.href='adminPriceCod.asp?codType='+this.value">
			<option value="L"<% If codType = "L" Then %> selected<% End If %>>
			<%=getadminPriceCodLngStr("DtxtLow")%></option>
			<option value="M"<% If codType = "M" Then %> selected<% End If %>>
			<%=getadminPriceCodLngStr("DtxtMedium")%></option>
			<option value="H"<% If codType = "H" Then %> selected<% End If %>>
			<%=getadminPriceCodLngStr("DtxtHigh")%></option>
			</select></td>
        </tr>
        	<tr>
          <td width="24" height="1" align="center" bgcolor="#E8F5FD"><b>
			<font face="Verdana" size="1" color="#3F7B96">.</font></b></td>
          <td width="51" height="1" align="center" bgcolor="#F5FBFE">
          <input type="text" id="FldCod" name="Field_11" size="5" style="font-family: Verdana; font-size: 10px; color: #3F7B96; border: 1px solid #68A6C0; background-color: #D9F0FD" value="<%=Server.HTMLEncode(rs("."))%>" onfocus="this.select()" maxlength="1"></td>
          <td width="28" height="1" align="center" bgcolor="#E8F5FD" class="style1">
			<strong>-</strong></td>
          <td width="50" height="1" bgcolor="#F5FBFE">
          <p align="center">
          <input type="text" id="FldCod" name="Field_" size="5" style="font-family: Verdana; font-size: 10px; color: #3F7B96; border: 1px solid #68A6C0; background-color: #D9F0FD" value="<%=Server.HTMLEncode(rs("-"))%>" onfocus="this.select()" maxlength="1"></td>
        	</tr>
        <tr>
          <td width="24" height="19" align="center" bgcolor="#E8F5FD"><b>
          <font size="1" face="Verdana" color="#3F7B96">1</font></b></td>
          <td width="51" height="19" align="center" bgcolor="#F5FBFE">
          <input type="text" id="FldCod" name="Field1" size="5" style="font-family: Verdana; font-size: 10px; color: #3F7B96; border: 1px solid #68A6C0; background-color: #D9F0FD" value="<%=Server.HTMLEncode(rs("1"))%>" onfocus="this.select()" maxlength="1"></td>
          <td width="28" height="19" align="center" bgcolor="#E8F5FD"><b>
			<font face="Verdana" size="1" color="#3F7B96">2</font></b></td>
          <td width="50" height="19" bgcolor="#F5FBFE">
          <p align="center">
          <input type="text" id="FldCod" name="Field2" size="5" style="font-family: Verdana; font-size: 10px; color: #3F7B96; border: 1px solid #68A6C0; background-color: #D9F0FD" value="<%=Server.HTMLEncode(rs("2"))%>" onfocus="this.select()" maxlength="1"></td>
        </tr>
        <tr>
          <td width="24" height="19" align="center" bgcolor="#E8F5FD"><b>
			<font face="Verdana" size="1" color="#3F7B96">3</font></b></td>
          <td width="51" height="19" align="center" bgcolor="#F5FBFE">
          <input name="Field3" id="FldCod" size="5" style="font-family: Verdana; font-size: 10px; color: #3F7B96; border: 1px solid #68A6C0; background-color: #D9F0FD" value="<%=Server.HTMLEncode(rs("3"))%>" onfocus="this.select()" maxlength="1"></td>
          <td width="28" height="19" align="center" bgcolor="#E8F5FD"><b>
			<font face="Verdana" size="1" color="#3F7B96">4</font></b></td>
          <td width="50" height="19" bgcolor="#F5FBFE">
          <p align="center">
          <input type="text" id="FldCod" name="Field4" size="5" style="font-family: Verdana; font-size: 10px; color: #3F7B96; border: 1px solid #68A6C0; background-color: #D9F0FD" value="<%=Server.HTMLEncode(rs("4"))%>" onfocus="this.select()" maxlength="1"></td>
        </tr>
        <tr>
          <td width="24" height="19" align="center" bgcolor="#E8F5FD"><b>
			<font face="Verdana" size="1" color="#3F7B96">5</font></b></td>
          <td width="51" height="19" align="center" bgcolor="#F5FBFE">
          <input type="text" id="FldCod" name="Field5" size="5" style="font-family: Verdana; font-size: 10px; color: #3F7B96; border: 1px solid #68A6C0; background-color: #D9F0FD" value="<%=Server.HTMLEncode(rs("5"))%>" onfocus="this.select()" maxlength="1"></td>
          <td width="28" height="19" align="center" bgcolor="#E8F5FD"><b>
			<font face="Verdana" size="1" color="#3F7B96">6</font></b></td>
          <td width="50" height="19" bgcolor="#F5FBFE">
          <p align="center">
          <input type="text" id="FldCod" name="Field6" size="5" style="font-family: Verdana; font-size: 10px; color: #3F7B96; border: 1px solid #68A6C0; background-color: #D9F0FD" value="<%=Server.HTMLEncode(rs("6"))%>" onfocus="this.select()" maxlength="1"></td>
        </tr>
        <tr>
          <td width="24" height="19" align="center" bgcolor="#E8F5FD"><b>
			<font face="Verdana" size="1" color="#3F7B96">7</font></b></td>
          <td width="51" height="19" align="center" bgcolor="#F5FBFE">
          <input type="text" id="FldCod" name="Field7" size="5" style="font-family: Verdana; font-size: 10px; color: #3F7B96; border: 1px solid #68A6C0; background-color: #D9F0FD" value="<%=Server.HTMLEncode(rs("7"))%>" onfocus="this.select()" maxlength="1"></td>
          <td width="28" height="19" align="center" bgcolor="#E8F5FD"><b>
			<font face="Verdana" size="1" color="#3F7B96">8</font></b></td>
          <td width="50" height="19" bgcolor="#F5FBFE">
          <p align="center">
          <input type="text" id="FldCod" name="Field8" size="5" style="font-family: Verdana; font-size: 10px; color: #3F7B96; border: 1px solid #68A6C0; background-color: #D9F0FD" value="<%=Server.HTMLEncode(rs("8"))%>" onfocus="this.select()" maxlength="1"></td>
        </tr>
        <tr>
          <td width="24" height="1" align="center" bgcolor="#E8F5FD"><b>
			<font face="Verdana" size="1" color="#3F7B96">9</font></b></td>
          <td width="51" height="1" align="center" bgcolor="#F5FBFE">
          <input type="text" id="FldCod" name="Field9" size="5" style="font-family: Verdana; font-size: 10px; color: #3F7B96; border: 1px solid #68A6C0; background-color: #D9F0FD" value="<%=Server.HTMLEncode(rs("9"))%>" onfocus="this.select()" maxlength="1"></td>
          <td width="28" height="1" align="center" bgcolor="#E8F5FD"><b>
			<font face="Verdana" size="1" color="#3F7B96">0</font></b></td>
          <td width="50" height="1" bgcolor="#F5FBFE">
          <p align="center">
          <input type="text" id="FldCod" name="Field0" size="5" style="font-family: Verdana; font-size: 10px; color: #3F7B96; border: 1px solid #68A6C0; background-color: #D9F0FD" value="<%=Server.HTMLEncode(rs("0"))%>" onfocus="this.select()" maxlength="1"></td>
        </tr>
        <tr>
          <td width="153" height="19" bgcolor="#F5FBFE" colspan="4">
          <p align="center">
          <input style="font-weight: bold; font-size: 10px; width: 75px; color: #68a6c0; font-family: Tahoma; height: 23px; border: 1px solid #68a6c0; background-color: #e5f1ff" type="submit" value="<%=getadminPriceCodLngStr("DtxtSave")%>" name="B1">
          </td>
        </tr>
      	<input type="hidden" name="submitCmd" value="adminPriceCod">
      	</form>
      </table>
        </div>
        
<!--#include file="bottom.asp" -->
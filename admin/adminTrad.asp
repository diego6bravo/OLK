<html <% If Session("rtl") <> "" Then %>dir="rtl" <% End If %>>

<!--#include file="chkLogin.asp" -->
<!--#include file="myHTMLEncode.asp"-->
<!--#include file="lang/adminTrad.asp" -->
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<% 
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

Dim oFCKeditor %>
<% set rs = Server.CreateObject("ADODB.RecordSet")

If Request("IsNew") <> "Y" Then
	sql = "select T0.LanID, T1." & Request("ColumnName") & " " & _
			"from OLKCommon..OLKLang T0 " & _
			"left outer join OLK" & Request("Table") & "AlterNames T1 on T1.LanID = T0.LanID "
	arrCol = Split(Request("ColumnID"), ",")
	arrVal = Split(Request("ID"), ",")
	For i = 0 to UBound(arrCol)
		If IsNumeric(arrVal(i)) Then
			sql = sql & "and T1." & arrCol(i) & " = " & arrVal(i) & " "
		Else
			sql = sql & "and T1." & arrCol(i) & " = N'" & arrVal(i) & "' "
		End If
	Next
Else
	If Request("NewValue") = "" Then
		sql = "select T0.LanID, '' " & Request("ColumnName") & " " & _
				"from OLKCommon..OLKLang T0 "
	Else
		arrValues = Split(Request("NewValue"), "{/}")
	End If
End If

If Request("NewValue") = "" Then rs.open sql, conn, 3, 1
%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-<%=getadminTradLngStr("charset")%>" />
<title><%=getadminTradLngStr("LttlTrans")%></title>
<style type="text/css">
body{
scrollbar-base-color:#336699;
scrollbar-highlight-color:#E1F1FF;
}

</style>
<link rel="stylesheet" type="text/css" href="style/style_pop.css"/>
<!-- #INCLUDE file="FCKeditor/fckeditor.asp" -->
</head>

<body onunload="opener.clearWin();" style="margin: 0;">
<script type="text/javascript">
var arrLng = '<% For i = 0 to UBound(myLanIndex) %><% If i > 0 Then %>, <% End If %><%=myLanIndex(i)(4)%><% Next %>'.split(', ');
function changeTradLnd(LanID)
{
	for (var i = 0;i<arrLng.length;i++)
	{
		document.getElementById('td' + arrLng[i]).style.display = arrLng[i] == LanID ? '' : 'none';
	}
}
</script>
<form method="post" action="<% If Request("IsNew") <> "Y" Then %>adminSubmit.asp<% Else %>adminTradReturn.asp<% End If %>">
	<table style="width: 100%" cellpadding="0">
		<tr class="TblGreenTlt">
			<td colspan="2">&nbsp;<%=getadminTradLngStr("LttlTrans")%></td>
		</tr>
		<tr class="TblGreenNrm">
			<td><%=getadminTradLngStr("LtxtLang")%></td>
			<td><select size="1" name="LanID" onchange="javascript:changeTradLnd(this.value);">
			<option></option>
			<% For i = 0 to UBound(myLanIndex) %>
			<option value="<%=myLanIndex(i)(4)%>"><%=myLanIndex(i)(1)%></option>
			<% Next %>
			</select></td>
		</tr>
		<% 
		If Request("NewValue") = "" Then
		do while not rs.eof %>
		<tr class="TblGreenNrm" id="td<%=rs("LanID")%>" style="display: none;">
			<td colspan="2" align="center">
			<% Select Case Request("Type")
				Case "T" %>
			<input name="txt<%=rs("LanID")%>" type="text" style="width: 98%" value="<% If Not IsNull(rs(1)) Then %><%=Server.HTMLEncode(rs(1))%><% End If %>" <% If rs("LanID") = 3 Then %>dir="rtl"<% End If %> />
			<%	Case "M" %>
			<textarea name="txt<%=rs("LanID")%>" rows="4" style="width: 98%" <% If rs("LanID") = 3 Then %>dir="rtl"<% End If %>><% If Not IsNull(rs(1)) Then %><%=Server.HTMLEncode(rs(1))%><% End If %></textarea>
			<% Case "R"
			Set oFCKeditor = New FCKeditor
			oFCKeditor.BasePath = "FCKeditor/"
			oFCKeditor.Height = 400
			oFCKEditor.ToolbarSet = "Custom"
			If Not IsNull(rs(1)) Then oFCKEditor.Value = myHTMLEncode(rs(1))
			oFCKEditor.Config("AutoDetectLanguage") = False
			If Session("myLng") <> "pt" Then
				oFCKEditor.Config("DefaultLanguage") = Session("myLng")
			Else
				oFCKEditor.Config("DefaultLanguage") = "pt-br"
			End If
			oFCKeditor.Create "txt" & rs("LanID")
			  End Select %>
			</td>
		</tr>
		<% rs.movenext
		loop
		Else
		For i = 0 to UBound(arrValues)
		LanID = Split(arrValues(i), "{=}")(0)
		Value = Split(arrValues(i), "{=}")(1) %>
		<tr class="TblGreenNrm" id="td<%=LanID%>" style="display: none;">
			<td colspan="2" align="center"><%
			 Select Case Request("Type")
				Case "T" 
			%><input name="txt<%=LanID%>" type="text" style="width: 98%" value="<%=myHTMLEncode(Value)%>" /><%
				Case "M" 
			%><textarea name="txt<%=LanID%>" rows="4" style="width: 98%"><%=myHTMLEncode(Value)%></textarea><%
			   Case "R"
			Set oFCKeditor = New FCKeditor
			oFCKeditor.BasePath = "FCKeditor/"
			oFCKeditor.Height = 400
			oFCKEditor.ToolbarSet = "Custom"
			oFCKEditor.Value = Value
			oFCKEditor.Config("AutoDetectLanguage") = False
			If Session("myLng") <> "pt" Then
				oFCKEditor.Config("DefaultLanguage") = Session("myLng")
			Else
				oFCKEditor.Config("DefaultLanguage") = "pt-br"
			End If
			oFCKeditor.Create "txt" & LanID
			  End Select %></td>
		</tr>
		<% Next
		End If %>
		<tr>
			<td colspan="2">
			<div align="center">
				<table border="0" cellpadding="0" id="table6" style="width: 100%">
					<tr>
						<td style="width: 75px">
						<p align="center">
						<input type="submit" value="<%=getadminTradLngStr("DtxtAccept")%>" name="btnAccept" class="OlkBtn" /></p>
						</td>
						<td >
							<hr color="#0D85C6" size="1"/></td>
						<td style="width: 75px">
						<p align="center">
						<input type="button" value="<%=getadminTradLngStr("DtxtCancel")%>" name="btnCancel" class="OlkBtn" onClick="javascript:window.close();" /></p></td>
					</tr>
				</table>
			</div>
			</td>
		</tr>
	</table>
	<input type="hidden" name="ColumnID" value='<%=Request("ColumnID")%>' />
	<input type="hidden" name="ColumnName" value='<%=Request("ColumnName")%>' />
	<input type="hidden" name="ID" value='<%=Request("ID")%>' />
	<input type="hidden" name="pop" value="Y" />
	<input type="hidden" name="Table" value='<%=Request("Table")%>' />
	<input type="hidden" name="Type" value='<%=Request("Type")%>' />
	<input type="hidden" name="submitCmd" value="adminTrad" />
	<input type="hidden" name="new" value='<%=Request("new")%>' />
</form>

</body>

</html>

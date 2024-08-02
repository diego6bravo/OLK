<html <% If Session("rtl") <> "" Then %>dir="rtl" <% End If %>>
<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="../chkLogin.asp" -->
<!--#include file="../myHTMLEncode.asp"-->
<!--#include file="lang/adminTrad.asp" -->

<%
Dim oFCKeditor

set rs = Server.CreateObject("ADODB.RecordSet")

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
<link rel="stylesheet" type="text/css" href="../design/0/style/stylePopUp.css" />
<!-- #INCLUDE file="../FCKeditor/fckeditor.asp" -->
</head>

<body onunload="opener.clearWin();" style="margin: 0; background-color: #F5FBFE; ">
<script type="text/javascript">
var arrLng = '<% For i = 0 to UBound(myLanIndex) %><% If i > 0 Then %>, <% End If %><%=myLanIndex(i)(4)%><% Next %>'.split(', ');
function changeTradLnd(LanID)
{
	for (var i = 0;i<arrLng.length;i++)
	{
		document.getElementById('td' + arrLng[i]).style.display = arrLng[i] == LanID ? '' : 'none';
	}
}
/*function acceptNew()
{
	var retVal = '';
	for (var i = 0;i<arrLng.length;i++)
	{
		if (i > 0) retVal += '{/}';
		retVal += arrLng[i] + '{=}' + document.getElementById('txt' + arrLng[i]).value;
	}
	opener.setNewFldTrad(retVal);
	window.close();
}*/
</script>
<form method="post" action="<% If Request("IsNew") <> "Y" Then %>adminTradSubmit.asp<% Else %>adminTradReturn.asp<% End If %>">
	<table style="width: 100%" cellpadding="0">
		<tr class="GeneralTblBold2">
			<td colspan="2">&nbsp;<%=getadminTradLngStr("LttlTrans")%></td>
		</tr>
		<tr class="GeneralTbl">
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
		<tr class="GeneralTbl" id="td<%=rs("LanID")%>" style="display: none;">
			<td colspan="2" align="center">
			<% Select Case Request("Type")
				Case "T" %>
			<input name="txt<%=rs("LanID")%>" type="text" style="width: 98%" value="<%=myHTMLEncode(rs(1))%>" />
			<%	Case "M" %>
			<textarea name="txt<%=rs("LanID")%>" rows="4" style="width: 98%"><%=myHTMLEncode(rs(1))%></textarea>
			<% Case "R"
			Set oFCKeditor = New FCKeditor
			oFCKeditor.BasePath = "../FCKeditor/"
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
		<tr class="GeneralTbl" id="td<%=LanID%>" style="display: none;">
			<td colspan="2" align="center">
			<% Select Case Request("Type")
				Case "T" %>
			<input name="txt<%=LanID%>" type="text" style="width: 98%" value="<%=myHTMLEncode(Value)%>" />
			<%	Case "M" %>
			<textarea name="txt<%=LanID%>" rows="4" style="width: 98%"><%=myHTMLEncode(Value)%></textarea>
			<% Case "R"
			Set oFCKeditor = New FCKeditor
			oFCKeditor.BasePath = "../FCKeditor/"
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
			  End Select %>
			</td>
		</tr>
		<% Next
		End If %>
		<tr class="GeneralTbl">
			<td><input name="btnAccept" type="submit" value="<%=getadminTradLngStr("DtxtAccept")%>" />&nbsp;</td>
			<td align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>">
			<input name="btnCancel" type="button" value="<%=getadminTradLngStr("DtxtCancel")%>" onclick="window.close();" />&nbsp;</td>
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

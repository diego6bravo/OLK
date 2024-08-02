<%@ Language=VBScript %> 
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title></title>
</head>
<body>
<% If session("OLKDB") = "" Then response.redirect "../lock.asp" %>
<!--#include file="../myHTMLEncode.asp"-->
<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<!--#include file="../lcidReturn.inc"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

set rs = Server.CreateObject("ADODB.Recordset")
sql = "select (select SDKID from r3_obscommon..tcif where companydb = N'" & Session("OlkDB") & "') As SDKId, " & _
	 "ItemCode, Quantity, WhsCode, UseBaseUn from r3_obscommon..doc1 where lognum = " & Session("RetVal") & " and " & _
	 "LineNum = " & Request("LineNum")
set rs = conn.execute(Sql)
SDKID = rs("SDKID")

If Request("WhsCode") = Rs("WhsCode") Then
	doSave()
	conn.close
	response.redirect "../operaciones.asp?cmd=cart"
Else
	set rd = Server.CreateObject("ADODB.RecordSet")
	sql = "select OLKCommon.dbo.DBOLKItemInv" & Session("ID") & "(N'" & RS("ItemCode") & "', N'" & saveHTMLDecode(Request("WhsCode"), False) & "', " & RS("Quantity") & ", N'" & Session("olkdb") & "', " & Session("RetVal") & ", " & Request("LineNum") & ") Verfy "
	set rd = conn.execute(sql)
	
	If rd("Verfy") = "Y" Then
		doSave()
	Else
		DoBack()
	End If
	
	conn.close
End If

Sub doSave
	If Request("LineMemo") <> "" Then LineMemo = "N'" & saveHTMLDecode(Request("LineMemo"), False) & "'" Else LineMemo = "NULL"
	If Request("VatGroup") <> "" Then VatGroup = "N'" & saveHTMLDecode(Request("VatGroup"), False) & "'" Else VatGroup = "NULL" 
	If Request("TaxCode") <> "" Then TaxCode = "N'" & saveHTMLDecode(Request("TaxCode"), False) & "'" Else TaxCode = "NULL"      
	
	sql = "select AliasID, TypeID, '" & SDKID & "'+AliasID As InsertID " & _
		  "from cufd T0 " & _
		  "left outer join OLKCUFD T1 on T1.TableID = T0.TableID and T1.FieldID = T0.FieldID " & _
		  "where T0.TableId = 'INV1' and AType in ('" & userType & "','T') and OP in ('T','P') and Active = 'Y'"
	set rs = server.createobject("ADODB.RecordSet")
	rs.open sql, conn, 3, 1
	
	If myApp.UnEmbPriceSet and Request("SaleType") = 3 Then
		Price = CDbl(Request("Price"))/CDbl(Request("SalPackUn"))
	Else
		Price = CDbl(Request("Price"))
	End If
	
	sql = "update r3_obscommon..doc1 set WhsCode = N'" & saveHTMLDecode(Request("WhsCode"), False) & "', VatGroup = " & VatGroup & ", TaxCode = " & TaxCode & ", Price = " & getNumeric(Price) & ", DiscPrcnt = " & getNumeric(Request("Discount")) & " "
	If Not rs.eof then
		do while not rs.eof
			strVal = saveHTMLDecode(Request("U_" & rs("AliasID")), False)
			If Request("U_" & rs("AliasID")) <> "" Then 
				Select Case rs("TypeID") 
					Case "D" 
						AliasVal = "Convert(datetime,'" & SaveSqlDate(strVal) & "',120)"
					Case "B"
						AliasVal = getNumeric(strVal)
					Case Else
						AliasVal = "N'" & strVal & "'" 
				End Select
			Else 
				AliasVal = "NULL"
			End If
			sql = sql & ", " & rs("InsertID") & " = " & AliasVal
		rs.movenext
		loop
	End If
	
	If myApp.SVer < 8 Then
		sql = sql & ", " & SDKID & "LineMemo = " & LineMemo
	End If
	
	sql = sql & " where lognum = " & Session("RetVal") & " and LineNum = " & Request("LineNum")
	conn.execute(sql)
	
	If myApp.SVer >= 8 Then
		sql = "declare @LogNum int set @LogNum = " & Session("RetVal") & " declare @LineNum int set @LineNum = " & Request("LineNum") & " "
		If Request("LineMemo") <> "" Then
			sql = sql & "if not exists(select '' from R3_OBsCommon..DOC10 where LogNum = @LogNum and LineType = 'T' and AfterLine = @LineNum) begin " & _
						"	insert R3_ObsCommon..DOC10(LogNum, LineNum, LineType, AfterLine, LineText) " & _
						"	values(@LogNum, @LineNum, 'T', @LineNum, " & LineMemo & ") " & _
						"end else begin " & _
						"	update R3_ObsCommon..DOC10 set LineText = " & LineMemo & " where LogNum = @LogNum and LineType = 'T' and AfterLine = @LineNum " & _
						"end "
		Else
			sql = sql & "delete R3_ObsCommon..DOC10 where LogNum = @LogNum and LineType = 'T' and AfterLine = @LineNum"
		End If
		conn.execute(sql)
	End If
End Sub

Sub DoBack %>
<form name="frm" method="post" action="../operaciones.asp">
<% For each itm in Request.Form
If itm <> "cmd" Then %>
<input type="hidden" name="<%=itm%>" value="<%=Request.Form(itm)%>">
<% 
End If
Next %>
<input type="hidden" name="cmd" value="cartEditLine">
<input type="hidden" name="err" value="Y">
</form>
<script language="javascript">document.frm.submit();</script>
<% End Sub %>
</body>
</html>
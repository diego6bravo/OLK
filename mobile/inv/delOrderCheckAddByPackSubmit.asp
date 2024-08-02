<%@ Language=VBScript %>
<!--#include file="../myHTMLEncode.asp"-->
<% If session("OLKDB") = "" Then response.redirect "../lock.asp" %>
<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

set rs = Server.CreateObject("ADODB.RecordSet")
sql = "select T1.ChkAllowOverload, T2.T1 oTable, T2.T2 oTable1, T1.DocType " & _
"from OLKCommon T0 " & _
"inner join OLKInOutSettings T1 on T1.ObjectCode = " & Session("ObjCode") & " and T1.Type = '" & Session("Type") & "' " & _
"inner join OLKDocConf T2 on T2.ObjectCode = T1.ObjectCode "
set rs = conn.execute(sql)

AllowOverload = rs("ChkAllowOverload") = "Y"
oTable = rs("oTable")
oTable1 = rs("oTable1")
If rs("DocType") = "S" Then
	NumIn = "Sale"
	Pack = "Sal"
	UnitMsr = "Sal"
ElseIf rs("DocType") = "P" Then
	NumIn = "Buy"
	Pack = "Pur"
	UnitMsr = "Buy"
End If

If myApp.EnableCodeBarsQry Then
	sql = "declare @CodeBars nvarchar(50) set @CodeBars = N'" & saveHTMLDecode(Request("txtItem"), False) & "' "
	sql = sql & "set @CodeBars = (" & myApp.CodeBarsQry & ")"
	
	sql = sql & " select @CodeBars CodeBars"
	set rs = conn.execute(sql)
	strCodeBars = saveHTMLDecode(rs("CodeBars"), False)
Else
	strCodeBars = saveHTMLDecode(Request("txtItem"), False)
End If

rs.close

sql = 	"select T2.ItemCode, T2." & Pack & "PackUn PackUn, T2.NumIn" & NumIn & " NumIn " & _
		"from " & oTable1 & " T0 " & _
		"inner join " & oTable & " T1 on T1.DocEntry = T0.DocEntry and T1.DocNum = " & Request("txtOrderNum") & " " & _
		"inner join OITM T2 on T2.ItemCode = T0.ItemCode " & _
		"where T0.WhsCode = N'" & Session("bodega") & "' and T0.LineStatus = 'O' and " & _
		"(T2.ItemCode = N'" & saveHTMLDecode(Request("txtItem"), False) & "' or " & _
		"T2.CodeBars = N'" & strCodeBars & "' "

If myApp.EnableSearchItmSupp Then
	sql = sql & " or T2.SuppCatNum = N'" & saveHTMLDecode(Request("txtItem"), False) & "'"
End If

		
sql = sql & ") Group By T2.ItemCode, T2." & Pack & "PackUn, T2.NumIn" & NumIn
rs.open sql, conn, 3, 1

If Not rs.Eof Then
	ItemCode = rs("ItemCode")
	
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKAddIOCheck" & Session("ID")
	cmd.ActiveConnection = connCommon
	cmd.Parameters.Refresh
	cmd("@LogNum") = Session("IORetVal")
	cmd("@ObjectCode") = Session("ObjCode")
	cmd("@Type") = Session("Type")
	cmd("@DocNum") = Request("txtOrderNum")
	cmd("@WhsCode") = saveHTMLDecode(Session("bodega"), True)
	cmd("@ItemCode") = saveHTMLDecode(ItemCode, True)
	
	PackUnit = 0
	If Request("chkRepack") = "Y" Then PackUnit = 1
	Select Case Request("rdUnit")
		Case 1
			cmd("@Unit") = CDbl(Request("txtQty"))
			cmd("@SBUnit") = 0
			cmd("@PackUnit") = CDbl(PackUnit)
		Case 2
			cmd("@Unit") = 0
			cmd("@SBUnit") = CDbl(Request("txtQty"))
			cmd("@PackUnit") = CDbl(PackUnit)
		Case 3
			cmd("@Unit") = 0
			cmd("@SBUnit") = CDbl(Request("txtQty"))*CDbl(rs("PackUn"))
			cmd("@PackUnit") = CDbl(Request("txtQty"))
	End Select
	
	cmd.execute()
	
	ManSerNum = cmd("@ManSerNum") = "Y"

	If Not ManSerNum Then
		response.redirect "../operaciones.asp?cmd=invChkInOutAddByPack&txtOrderNum=" & Request("txtOrderNum") & "&confirm=" & Server.HTMLEncode(ItemCode)
	Else
		response.redirect "../operaciones.asp?cmd=invChkInOutCheckSerial&txtOrderNum=" & Request("txtOrderNum") & "&ItemCode=" & Server.HTMLEncode(ItemCode) & "&retAddPack=Y"
	End If
Else
	response.redirect "../operaciones.asp?cmd=invChkInOutAddByPack&txtOrderNum=" & Request("txtOrderNum") & "&ErrItm=" & Server.HTMLEncode(Request("txtItem"))
End If
%>
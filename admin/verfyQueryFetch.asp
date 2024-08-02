<!--#include file="chkLogin.asp"-->
<!--#include file="myHTMLEncode.asp"-->
<!--#include file="getColType.inc"-->
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

If Request("Query") <> "" Then
	For i = 0 to UBound(myLanIndex)
		myItm = myLanIndex(i)
		If myItm(4) = CStr(Session("LanID")) Then
			sql = "set language " & myItm(5)
			conn.execute(sql)
			Exit For
		End If
	Next
	sql = ""

	Select Case Request("Type")
		Case "ItemDescQry"
			sql = "declare @SlpCode int set @SlpCode = -1 select Case When ("
		Case "SmallCat"
			sql = " select ItemCode from OITM where ItemCode in ("
	End Select
	
	sql = sql & Request("Query")
	
	Select Case Request("Type")
		Case "ItemDescQry"
			sql = sql & ") Then 'Y' Else 'N' End Verfy from R3_ObsCommon..DOC1 inner join OITM on OITM.ItemCode = DOC1.ItemCode collate database_default where DOC1.LogNum = 1 and DOC1.LineNum = 1"
			sql = QueryFunctions(sql)
			sql = Replace(sql, "@ItemCode", "OITM.ItemCode")
		Case "SmallCat"
			sql = sql & ")"
			If Request("CatType") = "I" Then sql = Replace(sql, "@ItemCode", "N''")
			sql = Replace(sql, "@CardCode", "N''")
			sql = Replace(sql, "@dbName", "N''")
			sql = Replace(sql, "@LanID", "1")
			sql = QueryFunctions(sql)
	End Select

End If

If Request("Query") <> "" Then
	On Error Resume Next
	set rs = conn.execute(sql)
	If Err.Number <> 0 Then
		Response.Write Err.Description
	Else
		Response.Write "ok"
	End If
End If
%>
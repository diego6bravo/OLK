<%
Function getTaxCode
	getTaxCode = getItemTaxCode(Request("Item"))
End Function

Function getItemTaxCode(ByVal Item)
	set cmdGetTaxCode = Server.CreateObject("ADODB.Command")
	cmdGetTaxCode.ActiveConnection = connCommon
	cmdGetTaxCode.CommandType = &H0004
	cmdGetTaxCode.CommandText = "DBOLKGetTaxCode" & Session("ID")
	cmdGetTaxCode.Parameters.Refresh()
	cmdGetTaxCode("@LogNum") = Session("RetVal")
	cmdGetTaxCode("@CardCode") = Session("UserName")
	cmdGetTaxCode("@ItemCode") = Item
	cmdGetTaxCode("@LawsSet") = myApp.LawsSet
	set rTax = Server.CreateObject("ADODB.RecordSet")
	set rTax = cmdGetTaxCode.execute()
	If Not rTax.EOF Then
		getItemTaxCode = rTax(0)
	Else
		getItemTaxCode = "Disabled"
	End If

End Function

Function getExpTaxCode
	set cmdGetTaxCode = Server.CreateObject("ADODB.Command")
	cmdGetTaxCode.ActiveConnection = connCommon
	cmdGetTaxCode.CommandType = &H0004
	cmdGetTaxCode.CommandText = "DBOLKGetExpTaxCode" & Session("ID")
	cmdGetTaxCode.Parameters.Refresh()
	cmdGetTaxCode("@LogNum") = Session("RetVal")
	set rTax = Server.CreateObject("ADODB.RecordSet")
	set rTax = cmdGetTaxCode.execute()
	If Not rTax.EOF Then
		getExpTaxCode = rTax(0)
	Else
		getExpTaxCode = "Disabled"
	End If
End Function
%>
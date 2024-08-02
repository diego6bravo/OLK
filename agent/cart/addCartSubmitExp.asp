<%@ Language=VBScript%>
<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="../chkLogin.asp" -->
<%

set rs = Server.CreateObject("ADODB.RecordSet")

AddItem = True
TaxCode = ""
If myApp.LawsSet = "MX" or myApp.LawsSet = "CL" or myApp.LawsSet = "CR" or myApp.LawsSet = "GT" or myApp.LawsSet = "US" or myApp.LawsSet = "CA" Then
	If Request("TaxCode") <> "" Then
	  	TaxCode = Request("TaxCode")
	Else
	  	TaxCode = getExpTaxCode
	End If
	
	If TaxCode = "Disabled" Then 
		TaxCode = ", NULL"
	ElseIf TaxCode = "" Then
		errMsg = "&err=tax&expItem=Y&tItem=" & Request("Item") & "&document=" & Request("document") & "&page=" & Request("page")
		AddItem = False
	Else
		TaxCode = ", '" & TaxCode & "' "
	End If
End If
If AddItem Then
  sql = "EXEC OLKCommon..DBOLKCartAddExp" & Session("ID") & " " & Session("RetVal") & ", " & Request("Item") & TaxCode
  conn.execute(sql)
End If


response.redirect "../cart.asp?update=Y"

%><!--#include file="../itemFunctions.asp" -->
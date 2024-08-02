<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>

<!--#include file="chkLogin.asp" -->
<!--#include file="myHTMLEncode.asp" -->
<html>

<head>
</head>

<body>

<% 
set rs = Server.CreateObject("ADODB.RecordSet") %>
<!--#include file="repVars.inc" -->
<%
	showLinkIcon = "N"
		
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandText = "DBOLKAdminRSTotals" & Session("ID")
	cmd.CommandType = adCmdStoredProc
	cmd.Parameters.Refresh()
	cmd("@rsIndex") = Request("rsIndex")
	cmd("@colName") = saveHTMLDecode(Request("colName"), True)
	cmd("@linkType") = Request("linkType")
	If Request("linkObject") <> "" Then
		cmd("@linkObject") = Request("linkObject")
		showLinkIcon = "Y"
	End If
	If Request("linkObjectPocket") <> "" Then
		cmd("@linkObjectPocket") = Request("linkObjectPocket")
		showLinkIcon = "Y"
	End If
    Select Case Request("linkType") 
    	Case "L" 
    		cmd("@linkLink") = saveHTMLDecode(Request("linkLink"), True)
    	Case "F"
    		cmd("@linkLink") = saveHTMLDecode(Request("linkLink"), True)
    		cmd("@linkLinkPocket") = saveHTMLDecode(Request("linkLinkPocket"), True)
    End Select

    If Request("Popup") = "Y" Then linkPopup = "Y" Else linkPopup = "N"
    If Request("linkCat") = "Y" Then linkCat = "Y" Else linkCat = "N"
    cmd("@linkPopup") = linkPopup
    cmd("@linkCat") = linkCat
    cmd("@Action") = "L"
    cmd.execute()
    
    set cmd = Server.CreateObject("ADODB.Command")
    cmd.ActiveConnection = connCommon
    cmd.CommandText = "DBOLKAdminRSLinksVars" & Session("ID")
    cmd.CommandType = adCmdStoredProc
    
	varVar = Split(Request("varVar"), ", ")
	For i = 0 to UBound(varVar)
	    cmd.Parameters.Refresh()
	    cmd("@rsIndex") = Request("rsIndex")
	    cmd("@colName") = saveHTMLDecode(Request("colName"), True)
		valBy = Request("valBy" & varVar(i))
		If valBy = "F" Then
			If Request("valValueF" & varVar(i)) <> "" Then 
				cmd("@valValue") = saveHTMLDecode(Request("valValueF" & varVar(i)), True)
			End If
		ElseIf valBy = "V" Then
			If Request("varDataType" & varVar(i)) <> "datetime" Then
				If Request("valValueV" & varVar(i)) <> "" Then
					cmd("@valValue") = saveHTMLDecode(Request("valValueV" & varVar(i)), True)
				End If
			Else
				If Request("colValDat" & varVar(i)) <> "" Then
					cmd("@valValDat") = SaveSqlDate(Request("colValDat" & varVar(i)))
				End If
			End If
		End If
		cmd("@varId") = varVar(i)
		cmd("@valBy") = valBy
		cmd.execute()
	next
	
	linkCol = ""
	sql = "select valValue from OLKRSLinksVars where rsIndex = " & Request("rsIndex") & " and valBy = 'F'"
	set rs = conn.execute(sql)
	do while not rs.eof
		If Not IsNull(rs("valValue")) Then 
			If linkCol <> "" Then linkCol = linkCol & "{S}"
			linkCol = linkCol & Replace(Replace(rs("valValue")," ",""), ".", "")
		End If
	rs.movenext
	loop
conn.close %>
<script language="javascript">
opener.clearWin();
opener.updateLinkImg('<%=showLinkIcon%>', '<%=linkCol%>');
window.close();
</script>
</body>

</html>
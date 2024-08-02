<%@ Language=VBScript %>
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
<%
           sql = "update R3_ObsCommon..tlog set status = 'B' where lognum in (" & Request("chkDel") & ") " & _
				"update OLKUAFControl set Status = 'X', ConfirmDate = getdate(), ConfirmUserSign= " & Session("vendid") & " where ExecAt = 'D3' and ObjectEntry in (select LogNum from R3_ObsCommon..TLOG where LogNum in (" & Request("chkDel") & ") and Object <> 24) and Status in ('O', 'E') " & _
				"update OLKUAFControl set Status = 'X', ConfirmDate = getdate(), ConfirmUserSign= " & Session("vendid") & " where ExecAt = 'R2' and ObjectEntry in (select LogNum from R3_ObsCommon..TLOG where LogNum in (" & Request("chkDel") & ") and Object = 24) and Status in ('O', 'E') "
conn.execute(sql)
conn.close
Session("retval") = ""

Select Case Request("go2")
	Case "I"
		actStr = "searchOpenedItems.asp"
	Case "C"
		actStr = "searchOpenedCards.asp"
	Case "D"
		actStr = "searchOpenedDocs.asp"
	Case "A" 
		actStr = "searchOpenedActivities.asp"
	Case "S" 
		actStr = "searchOpenedSO.asp"
	Case "AC"
		actStr = "activeClient.asp"
End Select
%>
<html>
<body onload="//document.frmGo.submit();">
<form name="frmGo" action="../<%=actStr%>" method="post">
<% For each Item in Request.Form
If Item <> "retval" and Item <> "del" and Item <> "cCode" Then %>
<input type="hidden" name="<%=Item%>" value="<%=saveHTMLDecode(Request(Item), False)%>">
<% End If
Next
For each Item in Request.QueryString
If Item <> "retval" and Item <> "del" and Item <> "cCode" Then %>
<input type="hidden" name="<%=Item%>" value="<%=saveHTMLDecode(Request(Item), False)%>">
<% End If
Next %>
</form>
<script type="text/javascript">document.frmGo.submit();</script>
</body>
</html>
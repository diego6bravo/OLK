<% addLngPathStr = "searchInc/" %>
<!--#include file="lang/backToSearch.asp" -->


<% 
btnDesc = getbackToSearchLngStr("LtxtNewSearch")
addLinkStr = ""
Select Case searchCmd
	Case "searchOfertsX"
		backLink = "ofertsSearch.asp"
	Case "searchItemX"
		backLink = "openedItems.asp"
	Case "searchCardX"
		backLink = "openedCards.asp"
	Case "searchActX"
		backLink = "openedActivities.asp"
	Case "searchSOX"
		backLink = "openedSO.asp"
	Case "searchDocX"
		If strScriptName <> "activeClient.asp" Then
			backLink = "openedDocs.asp"
		Else
			backLink = "openedDocs.asp"
			addLinkStr = "&CardCodeFrom=" & Session("username") & "&CardCodeTo=" & Session("username")
			btnDesc = getbackToSearchLngStr("LsearchDocs")
		End If
End Select %>
<table border="0" cellpadding="0" width="93%" id="table1">
	<tr>
		<td>
		<input type="button" value="<%=btnDesc%>" name="btnNewSearch" style="color: #FFFFFF; font-family: Verdana; font-size: 7pt; border: 1px solid #FFFFFF; background-color: #0065CE; width:98; height:18" onclick="javascript:window.location.href='<%=backLink%><%=addLinkStr%>'"></td>
	</tr>
</table>
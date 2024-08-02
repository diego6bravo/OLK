<%@ Language=VBScript %>
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>

<!--#include file="myHTMLEncode.asp" -->
<!--#include file="adminTradSave.asp"-->
<!--#include file="repVars.inc" -->
<%
set rs = server.createobject("ADODB.RecordSet")
set cmd = Server.CreateObject("ADODB.Command")

Select Case Request("cmd")
	Case "grp"
		If Request("GroupID") <> "" Then
			arrGrp = Split(Request("GroupID"), ", ")
			For i = 0 to UBound(arrGrp)
				grpID = arrGrp(i)
				sql = "update OLKFooterGroups set GroupName = N'" & saveHTMLDecode(Request("GroupName" & grpID), False) & "' where GroupID = " & grpID
				conn.execute(sql)
			Next
		End If
		If Request("GroupName") <> "" Then
			sql = "declare @GroupID int set @GroupID = IsNull((select Max(GroupID)+1 from OLKFooterGroups), 0) " & _
			"select @GroupID GroupID insert OLKFooterGroups(GroupID, GroupName) values(@GroupID, N'" & saveHTMLDecode(Request("GroupName"), False) & "')"
			set rs = conn.execute(sql)
			
			If Request("GroupNameTrad") <> "" Then
				SaveNewTrad Request("GroupNameTrad"), "FooterGroups", "GroupID", "alterGroupName", rs(0)
			End If
		End If
	Case "remGrp"
		sql = "declare @GroupID int set @GroupID = " & Request("GroupID") & " " & _
			"delete OLKFooterGroups where GroupID = @GroupID " & _
			"delete OLKFooterGroupsAlterNames where GroupID = @GroupID " & _
			"delete OLKFooterGroupsLinks where GroupID = @GroupID"
		conn.execute(sql)
	Case "editGrp"
		GroupID = CInt(Request("GroupID"))
		sql = "delete OLKFooterGroupsLinks where GroupID = " & GroupID
		conn.execute(sql)
		
		If Request("SecID") <> "" Then
			arrSecID = Split(Request("SecID"), ", ")
			For i = 0 to UBound(arrSecID)
				secID = arrSecID(i)
				sql = "insert OLKFooterGroupsLinks(GroupID, SecID, OrderID) values(" & GroupID & ", " & secID & ", " & Request("OrderID" & secID) & ")"
				conn.execute(sql)
			Next
		End If
		
		If Request("btnApply") <> "" Then 
			Response.Redirect "adminFooter.asp?GroupID=" & GroupID
		End If
End select

Response.Redirect "adminFooter.asp"

%>

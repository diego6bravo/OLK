<!--#include file="clientInc.asp"-->
<% Select Case userType
Case "C" %><!--#include file="clientTop.asp"-->
<% Case "V"
If Not myAut.HasAuthorization(1) Then Response.Redirect "unauthorized.asp"
 %><!--#include file="agentTop.asp"-->
<% End Select %>
<% addLngPathStr = "" %>
<!--|P:LangLink|-->
<% 
If Not IsNumeric(Request("navIndex")) Then Response.Redirect "unauthorized.asp"
NavIndex = CInt(Request("navIndex"))
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetObjectData" & Session("ID")
cmd.Parameters.Refresh()
cmd("@ObjType") = "S"
cmd("@ObjID") = 18
cmd("@UserType") = userType
set ra = cmd.execute()
strContent = ra("ObjContent")
strContent = Replace(strContent, "{SelDes}", SelDes)
strContent = Replace(strContent, "{rtl}", Session("rtl"))
strContent = Replace(strContent, "{dbName}", Session("olkdb"))
strContent = Replace(strContent, "{ImgMaxSize}", ra("ImgMaxSize"))
byX = CInt(ra("SubCols"))

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetSubNavIndex" & Session("ID")
cmd.Parameters.Refresh()
cmd("@LanID") = Session("LanID")
cmd("@UserType") = userType
cmd("@NavIndex") = NavIndex
set rs = Server.CreateObject("ADODB.RecordSet")
rs.open cmd, , 3, 1

set ri = Server.CreateObject("ADODB.RecordSet")
varx = 0 %>
<table border="0" cellpadding="0" width="100%">
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" cellspacing="1">
			<tr>
				<% 
				do while not rs.eof
				If varx = byX Then
					Response.Write "</tr><tr>"
					varx = 0
				End If

				NavIndex = rs("NavIndex")
				NavTitle = rs("NavTitle")
				NavDesc = rs("NavDesc")
				NavImgType = rs("NavImgType")
				NavImg = rs("NavImg")
				NavImgQry = rs("NavImgQry")
				NavType = rs("NavType")
				CatType = rs("CatType")
				
				If NavImgType = "I" Then
					If IsNull(NavImg) Then NavImg = "n_a.gif"
				ElseIf NavImgType = "Q" Then
					set ri = conn.execute(CStr(NavImgQry))
					If Not ri.Eof Then
						If Not IsNull(ri(0)) Then NavImg = ri(0) Else NavImg = "n_a.gif"
					Else
						NavImg = "n_a.gif"
					End If
				End If %>
				<td width="<%=100/byX%>%" valign="top" style="padding: 3px">
				<%
				strLoop = strContent
				If NavType <> "Q" Then
					strLoop = Replace(strLoop, "{Link}", "subNavIndex.asp?navIndex=" & NavIndex) 
				Else
					strLoop = Replace(strLoop, "{Link}", "javascript:goNavQry(" & NavIndex & ", '" & CatType & "')") 
				End If
				strLoop = Replace(strLoop, "{NavTitle}", NavTitle)
				strLoop = Replace(strLoop, "{NavImg}", NavImg)
				strLoop = Replace(strLoop, "{NavDesc}", NavDesc)
				Response.Write strLoop
				%>
				</td>
				<% 
				varx = varx + 1
				rs.movenext
				loop %>
			</tr>
		</table>
		</td>
	</tr>
</table>
<% set ri = nothing %>
<% If userType = "C" Then %>
<script language="javascript">
function goNavQry(navIndex, CatType)
{
	var strNewFld = '<input type="hidden" name="navIndex" value="' + navIndex + '">';
	var newFld = document.createElement(strNewFld);
	document.formSmallSearch.appendChild(newFld);
	document.formSmallSearch.string.value = '';
	document.formSmallSearch.document.value = CatType;
	document.formSmallSearch.submit();
}
</script>
<% End If %>
<% Select Case userType
Case "C" %><!--#include file="clientBottom.asp"-->
<% Case "V" %><!--#include file="agentBottom.asp"-->
<% End Select %>
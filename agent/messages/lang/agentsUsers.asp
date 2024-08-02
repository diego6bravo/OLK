<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "agentsUsers.xml"
set docagentsUsers = server.CreateObject("MSXML2.DOMDocument")
docagentsUsers.async = False
DocagentsUsers.Load(server.MapPath(xmlfilename)) 
docagentsUsers.setProperty "SelectionLanguage", "XPath"
set selectedagentsUsersnode = docagentsUsers.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedagentsUsersnodes=docagentsUsers.documentElement.selectNodes("/languages/language")
function getagentsUsersLngStr(instring)
	temp = selectedagentsUsersnode.selectSingleNode(instring).text
	getagentsUsersLngStr = temp
end function
%>

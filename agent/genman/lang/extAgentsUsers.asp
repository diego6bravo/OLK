<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "extAgentsUsers.xml"
set docextAgentsUsers = server.CreateObject("MSXML2.DOMDocument")
docextAgentsUsers.async = False
DocextAgentsUsers.Load(server.MapPath(xmlfilename)) 
docextAgentsUsers.setProperty "SelectionLanguage", "XPath"
set selectedextAgentsUsersnode = docextAgentsUsers.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedextAgentsUsersnodes=docextAgentsUsers.documentElement.selectNodes("/languages/language")
function getextAgentsUsersLngStr(instring)
	temp = selectedextAgentsUsersnode.selectSingleNode(instring).text
	getextAgentsUsersLngStr = temp
end function
%>

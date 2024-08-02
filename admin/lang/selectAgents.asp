<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "selectAgents.xml"
set docselectAgents = server.CreateObject("MSXML2.DOMDocument")
docselectAgents.async = False
DocselectAgents.Load(server.MapPath(xmlfilename)) 
docselectAgents.setProperty "SelectionLanguage", "XPath"
set selectedselectAgentsnode = docselectAgents.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedselectAgentsnodes=docselectAgents.documentElement.selectNodes("/languages/language")
function getselectAgentsLngStr(instring)
	temp = selectedselectAgentsnode.selectSingleNode(instring).text
	getselectAgentsLngStr = temp
end function
%>

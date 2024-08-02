<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "agentBottom.xml"
set docagentBottom = server.CreateObject("MSXML2.DOMDocument")
docagentBottom.async = False
DocagentBottom.Load(server.MapPath(xmlfilename)) 
docagentBottom.setProperty "SelectionLanguage", "XPath"
set selectedagentBottomnode = docagentBottom.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedagentBottomnodes=docagentBottom.documentElement.selectNodes("/languages/language")
function getagentBottomLngStr(instring)
	temp = selectedagentBottomnode.selectSingleNode(instring).text
	getagentBottomLngStr = temp
end function
%>

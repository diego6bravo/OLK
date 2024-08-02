<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "agentClientSubmit.xml"
set docagentClientSubmit = server.CreateObject("MSXML2.DOMDocument")
docagentClientSubmit.async = False
DocagentClientSubmit.Load(server.MapPath(xmlfilename)) 
docagentClientSubmit.setProperty "SelectionLanguage", "XPath"
set selectedagentClientSubmitnode = docagentClientSubmit.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedagentClientSubmitnodes=docagentClientSubmit.documentElement.selectNodes("/languages/language")
function getagentClientSubmitLngStr(instring)
	temp = selectedagentClientSubmitnode.selectSingleNode(instring).text
	getagentClientSubmitLngStr = temp
end function
%>

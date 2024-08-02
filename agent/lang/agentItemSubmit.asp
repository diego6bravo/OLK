<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "agentItemSubmit.xml"
set docagentItemSubmit = server.CreateObject("MSXML2.DOMDocument")
docagentItemSubmit.async = False
DocagentItemSubmit.Load(server.MapPath(xmlfilename)) 
docagentItemSubmit.setProperty "SelectionLanguage", "XPath"
set selectedagentItemSubmitnode = docagentItemSubmit.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedagentItemSubmitnodes=docagentItemSubmit.documentElement.selectNodes("/languages/language")
function getagentItemSubmitLngStr(instring)
	temp = selectedagentItemSubmitnode.selectSingleNode(instring).text
	getagentItemSubmitLngStr = temp
end function
%>

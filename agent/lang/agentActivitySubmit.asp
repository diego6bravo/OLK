<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "agentActivitySubmit.xml"
set docagentActivitySubmit = server.CreateObject("MSXML2.DOMDocument")
docagentActivitySubmit.async = False
DocagentActivitySubmit.Load(server.MapPath(xmlfilename)) 
docagentActivitySubmit.setProperty "SelectionLanguage", "XPath"
set selectedagentActivitySubmitnode = docagentActivitySubmit.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedagentActivitySubmitnodes=docagentActivitySubmit.documentElement.selectNodes("/languages/language")
function getagentActivitySubmitLngStr(instring)
	temp = selectedagentActivitySubmitnode.selectSingleNode(instring).text
	getagentActivitySubmitLngStr = temp
end function
%>

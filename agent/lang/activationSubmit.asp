<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "activationSubmit.xml"
set docactivationSubmit = server.CreateObject("MSXML2.DOMDocument")
docactivationSubmit.async = False
DocactivationSubmit.Load(server.MapPath(xmlfilename)) 
docactivationSubmit.setProperty "SelectionLanguage", "XPath"
set selectedactivationSubmitnode = docactivationSubmit.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedactivationSubmitnodes=docactivationSubmit.documentElement.selectNodes("/languages/language")
function getactivationSubmitLngStr(instring)
	temp = selectedactivationSubmitnode.selectSingleNode(instring).text
	getactivationSubmitLngStr = temp
end function
%>

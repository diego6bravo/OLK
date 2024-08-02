<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "activation.xml"
set docactivation = server.CreateObject("MSXML2.DOMDocument")
docactivation.async = False
Docactivation.Load(server.MapPath(xmlfilename)) 
docactivation.setProperty "SelectionLanguage", "XPath"
set selectedactivationnode = docactivation.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedactivationnodes=docactivation.documentElement.selectNodes("/languages/language")
function getactivationLngStr(instring)
	temp = selectedactivationnode.selectSingleNode(instring).text
	getactivationLngStr = temp
end function
%>

<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "activationReason.xml"
set docactivationReason = server.CreateObject("MSXML2.DOMDocument")
docactivationReason.async = False
DocactivationReason.Load(server.MapPath(xmlfilename)) 
docactivationReason.setProperty "SelectionLanguage", "XPath"
set selectedactivationReasonnode = docactivationReason.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedactivationReasonnodes=docactivationReason.documentElement.selectNodes("/languages/language")
function getactivationReasonLngStr(instring)
	temp = selectedactivationReasonnode.selectSingleNode(instring).text
	getactivationReasonLngStr = temp
end function
%>

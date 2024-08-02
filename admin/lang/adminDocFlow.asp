<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminDocFlow.xml"
set docadminDocFlow = server.CreateObject("MSXML2.DOMDocument")
docadminDocFlow.async = False
DocadminDocFlow.Load(server.MapPath(xmlfilename)) 
docadminDocFlow.setProperty "SelectionLanguage", "XPath"
set selectedadminDocFlownode = docadminDocFlow.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminDocFlownodes=docadminDocFlow.documentElement.selectNodes("/languages/language")
function getadminDocFlowLngStr(instring)
	temp = selectedadminDocFlownode.selectSingleNode(instring).text
	getadminDocFlowLngStr = temp
end function
%>

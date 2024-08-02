<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminAgentsAutCopy.xml"
set docadminAgentsAutCopy = server.CreateObject("MSXML2.DOMDocument")
docadminAgentsAutCopy.async = False
DocadminAgentsAutCopy.Load(server.MapPath(xmlfilename)) 
docadminAgentsAutCopy.setProperty "SelectionLanguage", "XPath"
set selectedadminAgentsAutCopynode = docadminAgentsAutCopy.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminAgentsAutCopynodes=docadminAgentsAutCopy.documentElement.selectNodes("/languages/language")
function getadminAgentsAutCopyLngStr(instring)
	temp = selectedadminAgentsAutCopynode.selectSingleNode(instring).text
	getadminAgentsAutCopyLngStr = temp
end function
%>

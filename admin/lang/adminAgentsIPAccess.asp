<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminAgentsIPAccess.xml"
set docadminAgentsIPAccess = server.CreateObject("MSXML2.DOMDocument")
docadminAgentsIPAccess.async = False
DocadminAgentsIPAccess.Load(server.MapPath(xmlfilename)) 
docadminAgentsIPAccess.setProperty "SelectionLanguage", "XPath"
set selectedadminAgentsIPAccessnode = docadminAgentsIPAccess.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminAgentsIPAccessnodes=docadminAgentsIPAccess.documentElement.selectNodes("/languages/language")
function getadminAgentsIPAccessLngStr(instring)
	temp = selectedadminAgentsIPAccessnode.selectSingleNode(instring).text
	getadminAgentsIPAccessLngStr = temp
end function
%>

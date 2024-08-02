<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminAgentsAccess.xml"
set docadminAgentsAccess = server.CreateObject("MSXML2.DOMDocument")
docadminAgentsAccess.async = False
DocadminAgentsAccess.Load(server.MapPath(xmlfilename)) 
docadminAgentsAccess.setProperty "SelectionLanguage", "XPath"
set selectedadminAgentsAccessnode = docadminAgentsAccess.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminAgentsAccessnodes=docadminAgentsAccess.documentElement.selectNodes("/languages/language")
function getadminAgentsAccessLngStr(instring)
	temp = selectedadminAgentsAccessnode.selectSingleNode(instring).text
	getadminAgentsAccessLngStr = temp
end function
%>

<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminAgentsRGAccess.xml"
set docadminAgentsRGAccess = server.CreateObject("MSXML2.DOMDocument")
docadminAgentsRGAccess.async = False
DocadminAgentsRGAccess.Load(server.MapPath(xmlfilename)) 
docadminAgentsRGAccess.setProperty "SelectionLanguage", "XPath"
set selectedadminAgentsRGAccessnode = docadminAgentsRGAccess.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminAgentsRGAccessnodes=docadminAgentsRGAccess.documentElement.selectNodes("/languages/language")
function getadminAgentsRGAccessLngStr(instring)
	temp = selectedadminAgentsRGAccessnode.selectSingleNode(instring).text
	getadminAgentsRGAccessLngStr = temp
end function
%>

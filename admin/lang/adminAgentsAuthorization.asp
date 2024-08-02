<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminAgentsAuthorization.xml"
set docadminAgentsAuthorization = server.CreateObject("MSXML2.DOMDocument")
docadminAgentsAuthorization.async = False
DocadminAgentsAuthorization.Load(server.MapPath(xmlfilename)) 
docadminAgentsAuthorization.setProperty "SelectionLanguage", "XPath"
set selectedadminAgentsAuthorizationnode = docadminAgentsAuthorization.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminAgentsAuthorizationnodes=docadminAgentsAuthorization.documentElement.selectNodes("/languages/language")
function getadminAgentsAuthorizationLngStr(instring)
	temp = selectedadminAgentsAuthorizationnode.selectSingleNode(instring).text
	getadminAgentsAuthorizationLngStr = temp
end function
%>

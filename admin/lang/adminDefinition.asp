<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminDefinition.xml"
set docadminDefinition = server.CreateObject("MSXML2.DOMDocument")
docadminDefinition.async = False
DocadminDefinition.Load(server.MapPath(xmlfilename)) 
docadminDefinition.setProperty "SelectionLanguage", "XPath"
set selectedadminDefinitionnode = docadminDefinition.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminDefinitionnodes=docadminDefinition.documentElement.selectNodes("/languages/language")
function getadminDefinitionLngStr(instring)
	temp = selectedadminDefinitionnode.selectSingleNode(instring).text
	getadminDefinitionLngStr = temp
end function
%>

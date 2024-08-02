<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminSystem.xml"
set docadminSystem = server.CreateObject("MSXML2.DOMDocument")
docadminSystem.async = False
DocadminSystem.Load(server.MapPath(xmlfilename)) 
docadminSystem.setProperty "SelectionLanguage", "XPath"
set selectedadminSystemnode = docadminSystem.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminSystemnodes=docadminSystem.documentElement.selectNodes("/languages/language")
function getadminSystemLngStr(instring)
	temp = selectedadminSystemnode.selectSingleNode(instring).text
	getadminSystemLngStr = temp
end function
%>

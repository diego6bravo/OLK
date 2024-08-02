<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminSingleAccess.xml"
set docadminSingleAccess = server.CreateObject("MSXML2.DOMDocument")
docadminSingleAccess.async = False
DocadminSingleAccess.Load(server.MapPath(xmlfilename)) 
docadminSingleAccess.setProperty "SelectionLanguage", "XPath"
set selectedadminSingleAccessnode = docadminSingleAccess.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminSingleAccessnodes=docadminSingleAccess.documentElement.selectNodes("/languages/language")
function getadminSingleAccessLngStr(instring)
	temp = selectedadminSingleAccessnode.selectSingleNode(instring).text
	getadminSingleAccessLngStr = temp
end function
%>

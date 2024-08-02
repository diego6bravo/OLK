<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminSingleIPAccess.xml"
set docadminSingleIPAccess = server.CreateObject("MSXML2.DOMDocument")
docadminSingleIPAccess.async = False
DocadminSingleIPAccess.Load(server.MapPath(xmlfilename)) 
docadminSingleIPAccess.setProperty "SelectionLanguage", "XPath"
set selectedadminSingleIPAccessnode = docadminSingleIPAccess.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminSingleIPAccessnodes=docadminSingleIPAccess.documentElement.selectNodes("/languages/language")
function getadminSingleIPAccessLngStr(instring)
	temp = selectedadminSingleIPAccessnode.selectSingleNode(instring).text
	getadminSingleIPAccessLngStr = temp
end function
%>

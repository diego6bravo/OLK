<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminSec.xml"
set docadminSec = server.CreateObject("MSXML2.DOMDocument")
docadminSec.async = False
DocadminSec.Load(server.MapPath(xmlfilename)) 
docadminSec.setProperty "SelectionLanguage", "XPath"
set selectedadminSecnode = docadminSec.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminSecnodes=docadminSec.documentElement.selectNodes("/languages/language")
function getadminSecLngStr(instring)
	temp = selectedadminSecnode.selectSingleNode(instring).text
	getadminSecLngStr = temp
end function
%>

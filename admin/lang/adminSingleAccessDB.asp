<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminSingleAccessDB.xml"
set docadminSingleAccessDB = server.CreateObject("MSXML2.DOMDocument")
docadminSingleAccessDB.async = False
DocadminSingleAccessDB.Load(server.MapPath(xmlfilename)) 
docadminSingleAccessDB.setProperty "SelectionLanguage", "XPath"
set selectedadminSingleAccessDBnode = docadminSingleAccessDB.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminSingleAccessDBnodes=docadminSingleAccessDB.documentElement.selectNodes("/languages/language")
function getadminSingleAccessDBLngStr(instring)
	temp = selectedadminSingleAccessDBnode.selectSingleNode(instring).text
	getadminSingleAccessDBLngStr = temp
end function
%>

<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminAlterNames.xml"
set docadminAlterNames = server.CreateObject("MSXML2.DOMDocument")
docadminAlterNames.async = False
DocadminAlterNames.Load(server.MapPath(xmlfilename)) 
docadminAlterNames.setProperty "SelectionLanguage", "XPath"
set selectedadminAlterNamesnode = docadminAlterNames.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminAlterNamesnodes=docadminAlterNames.documentElement.selectNodes("/languages/language")
function getadminAlterNamesLngStr(instring)
	temp = selectedadminAlterNamesnode.selectSingleNode(instring).text
	getadminAlterNamesLngStr = temp
end function
%>

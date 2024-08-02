<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminCustomSearchProp.xml"
set docadminCustomSearchProp = server.CreateObject("MSXML2.DOMDocument")
docadminCustomSearchProp.async = False
DocadminCustomSearchProp.Load(server.MapPath(xmlfilename)) 
docadminCustomSearchProp.setProperty "SelectionLanguage", "XPath"
set selectedadminCustomSearchPropnode = docadminCustomSearchProp.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminCustomSearchPropnodes=docadminCustomSearchProp.documentElement.selectNodes("/languages/language")
function getadminCustomSearchPropLngStr(instring)
	temp = selectedadminCustomSearchPropnode.selectSingleNode(instring).text
	getadminCustomSearchPropLngStr = temp
end function
%>

<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adSearchProp.xml"
set docadSearchProp = server.CreateObject("MSXML2.DOMDocument")
docadSearchProp.async = False
DocadSearchProp.Load(server.MapPath(xmlfilename)) 
docadSearchProp.setProperty "SelectionLanguage", "XPath"
set selectedadSearchPropnode = docadSearchProp.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadSearchPropnodes=docadSearchProp.documentElement.selectNodes("/languages/language")
function getadSearchPropLngStr(instring)
	temp = selectedadSearchPropnode.selectSingleNode(instring).text
	getadSearchPropLngStr = temp
end function
%>

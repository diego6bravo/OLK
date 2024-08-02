<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adCustomSearch.xml"
set docadCustomSearch = server.CreateObject("MSXML2.DOMDocument")
docadCustomSearch.async = False
DocadCustomSearch.Load(server.MapPath(xmlfilename)) 
docadCustomSearch.setProperty "SelectionLanguage", "XPath"
set selectedadCustomSearchnode = docadCustomSearch.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadCustomSearchnodes=docadCustomSearch.documentElement.selectNodes("/languages/language")
function getadCustomSearchLngStr(instring)
	temp = selectedadCustomSearchnode.selectSingleNode(instring).text
	getadCustomSearchLngStr = temp
end function
%>

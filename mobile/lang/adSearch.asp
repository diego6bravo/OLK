<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adSearch.xml"
set docadSearch = server.CreateObject("MSXML2.DOMDocument")
docadSearch.async = False
DocadSearch.Load(server.MapPath(xmlfilename)) 
docadSearch.setProperty "SelectionLanguage", "XPath"
set selectedadSearchnode = docadSearch.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadSearchnodes=docadSearch.documentElement.selectNodes("/languages/language")
function getadSearchLngStr(instring)
	temp = selectedadSearchnode.selectSingleNode(instring).text
	getadSearchLngStr = temp
end function
%>

<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminCustomSearch.xml"
set docadminCustomSearch = server.CreateObject("MSXML2.DOMDocument")
docadminCustomSearch.async = False
DocadminCustomSearch.Load(server.MapPath(xmlfilename)) 
docadminCustomSearch.setProperty "SelectionLanguage", "XPath"
set selectedadminCustomSearchnode = docadminCustomSearch.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminCustomSearchnodes=docadminCustomSearch.documentElement.selectNodes("/languages/language")
function getadminCustomSearchLngStr(instring)
	temp = selectedadminCustomSearchnode.selectSingleNode(instring).text
	getadminCustomSearchLngStr = temp
end function
%>

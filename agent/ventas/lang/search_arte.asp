<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "search_arte.xml"
set docsearch_arte = server.CreateObject("MSXML2.DOMDocument")
docsearch_arte.async = False
Docsearch_arte.Load(server.MapPath(xmlfilename)) 
docsearch_arte.setProperty "SelectionLanguage", "XPath"
set selectedsearch_artenode = docsearch_arte.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedsearch_artenodes=docsearch_arte.documentElement.selectNodes("/languages/language")
function getsearch_arteLngStr(instring)
	temp = selectedsearch_artenode.selectSingleNode(instring).text
	getsearch_arteLngStr = temp
end function
%>

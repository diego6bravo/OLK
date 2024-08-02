<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "search_itemX.xml"
set docsearch_itemX = server.CreateObject("MSXML2.DOMDocument")
docsearch_itemX.async = False
Docsearch_itemX.Load(server.MapPath(xmlfilename)) 
docsearch_itemX.setProperty "SelectionLanguage", "XPath"
set selectedsearch_itemXnode = docsearch_itemX.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedsearch_itemXnodes=docsearch_itemX.documentElement.selectNodes("/languages/language")
function getsearch_itemXLngStr(instring)
	temp = selectedsearch_itemXnode.selectSingleNode(instring).text
	getsearch_itemXLngStr = temp
end function
%>

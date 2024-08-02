<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "search_cardX.xml"
set docsearch_cardX = server.CreateObject("MSXML2.DOMDocument")
docsearch_cardX.async = False
Docsearch_cardX.Load(server.MapPath(xmlfilename)) 
docsearch_cardX.setProperty "SelectionLanguage", "XPath"
set selectedsearch_cardXnode = docsearch_cardX.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedsearch_cardXnodes=docsearch_cardX.documentElement.selectNodes("/languages/language")
function getsearch_cardXLngStr(instring)
	temp = selectedsearch_cardXnode.selectSingleNode(instring).text
	getsearch_cardXLngStr = temp
end function
%>

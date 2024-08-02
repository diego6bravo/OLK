<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "search_ventasX.xml"
set docsearch_ventasX = server.CreateObject("MSXML2.DOMDocument")
docsearch_ventasX.async = False
Docsearch_ventasX.Load(server.MapPath(xmlfilename)) 
docsearch_ventasX.setProperty "SelectionLanguage", "XPath"
set selectedsearch_ventasXnode = docsearch_ventasX.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedsearch_ventasXnodes=docsearch_ventasX.documentElement.selectNodes("/languages/language")
function getsearch_ventasXLngStr(instring)
	temp = selectedsearch_ventasXnode.selectSingleNode(instring).text
	getsearch_ventasXLngStr = temp
end function
%>

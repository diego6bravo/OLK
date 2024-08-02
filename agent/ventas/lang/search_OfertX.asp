<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "search_OfertX.xml"
set docsearch_OfertX = server.CreateObject("MSXML2.DOMDocument")
docsearch_OfertX.async = False
Docsearch_OfertX.Load(server.MapPath(xmlfilename)) 
docsearch_OfertX.setProperty "SelectionLanguage", "XPath"
set selectedsearch_OfertXnode = docsearch_OfertX.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedsearch_OfertXnodes=docsearch_OfertX.documentElement.selectNodes("/languages/language")
function getsearch_OfertXLngStr(instring)
	temp = selectedsearch_OfertXnode.selectSingleNode(instring).text
	getsearch_OfertXLngStr = temp
end function
%>

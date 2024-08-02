<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "search_inte.xml"
set docsearch_inte = server.CreateObject("MSXML2.DOMDocument")
docsearch_inte.async = False
Docsearch_inte.Load(server.MapPath(xmlfilename)) 
docsearch_inte.setProperty "SelectionLanguage", "XPath"
set selectedsearch_intenode = docsearch_inte.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedsearch_intenodes=docsearch_inte.documentElement.selectNodes("/languages/language")
function getsearch_inteLngStr(instring)
	temp = selectedsearch_intenode.selectSingleNode(instring).text
	getsearch_inteLngStr = temp
end function
%>

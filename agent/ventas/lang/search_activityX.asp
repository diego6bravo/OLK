<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "search_activityX.xml"
set docsearch_activityX = server.CreateObject("MSXML2.DOMDocument")
docsearch_activityX.async = False
Docsearch_activityX.Load(server.MapPath(xmlfilename)) 
docsearch_activityX.setProperty "SelectionLanguage", "XPath"
set selectedsearch_activityXnode = docsearch_activityX.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedsearch_activityXnodes=docsearch_activityX.documentElement.selectNodes("/languages/language")
function getsearch_activityXLngStr(instring)
	temp = selectedsearch_activityXnode.selectSingleNode(instring).text
	getsearch_activityXLngStr = temp
end function
%>

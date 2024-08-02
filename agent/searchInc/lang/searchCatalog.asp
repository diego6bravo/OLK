<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "searchCatalog.xml"
set docsearchCatalog = server.CreateObject("MSXML2.DOMDocument")
docsearchCatalog.async = False
DocsearchCatalog.Load(server.MapPath(xmlfilename)) 
docsearchCatalog.setProperty "SelectionLanguage", "XPath"
set selectedsearchCatalognode = docsearchCatalog.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedsearchCatalognodes=docsearchCatalog.documentElement.selectNodes("/languages/language")
function getsearchCatalogLngStr(instring)
	temp = selectedsearchCatalognode.selectSingleNode(instring).text
	getsearchCatalogLngStr = temp
end function
%>

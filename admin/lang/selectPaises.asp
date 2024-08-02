<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "selectPaises.xml"
set docselectPaises = server.CreateObject("MSXML2.DOMDocument")
docselectPaises.async = False
DocselectPaises.Load(server.MapPath(xmlfilename)) 
docselectPaises.setProperty "SelectionLanguage", "XPath"
set selectedselectPaisesnode = docselectPaises.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedselectPaisesnodes=docselectPaises.documentElement.selectNodes("/languages/language")
function getselectPaisesLngStr(instring)
	temp = selectedselectPaisesnode.selectSingleNode(instring).text
	getselectPaisesLngStr = temp
end function
%>

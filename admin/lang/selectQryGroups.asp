<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "selectQryGroups.xml"
set docselectQryGroups = server.CreateObject("MSXML2.DOMDocument")
docselectQryGroups.async = False
DocselectQryGroups.Load(server.MapPath(xmlfilename)) 
docselectQryGroups.setProperty "SelectionLanguage", "XPath"
set selectedselectQryGroupsnode = docselectQryGroups.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedselectQryGroupsnodes=docselectQryGroups.documentElement.selectNodes("/languages/language")
function getselectQryGroupsLngStr(instring)
	temp = selectedselectQryGroupsnode.selectSingleNode(instring).text
	getselectQryGroupsLngStr = temp
end function
%>

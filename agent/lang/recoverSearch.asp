<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "recoverSearch.xml"
set docrecoverSearch = server.CreateObject("MSXML2.DOMDocument")
docrecoverSearch.async = False
DocrecoverSearch.Load(server.MapPath(xmlfilename)) 
docrecoverSearch.setProperty "SelectionLanguage", "XPath"
set selectedrecoverSearchnode = docrecoverSearch.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedrecoverSearchnodes=docrecoverSearch.documentElement.selectNodes("/languages/language")
function getrecoverSearchLngStr(instring)
	temp = selectedrecoverSearchnode.selectSingleNode(instring).text
	getrecoverSearchLngStr = temp
end function
%>

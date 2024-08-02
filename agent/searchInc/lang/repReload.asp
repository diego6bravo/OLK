<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "repReload.xml"
set docrepReload = server.CreateObject("MSXML2.DOMDocument")
docrepReload.async = False
DocrepReload.Load(server.MapPath(xmlfilename)) 
docrepReload.setProperty "SelectionLanguage", "XPath"
set selectedrepReloadnode = docrepReload.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedrepReloadnodes=docrepReload.documentElement.selectNodes("/languages/language")
function getrepReloadLngStr(instring)
	temp = selectedrepReloadnode.selectSingleNode(instring).text
	getrepReloadLngStr = temp
end function
%>

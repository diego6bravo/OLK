<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "listpenSearch.xml"
set doclistpenSearch = server.CreateObject("MSXML2.DOMDocument")
doclistpenSearch.async = False
DoclistpenSearch.Load(server.MapPath(xmlfilename)) 
doclistpenSearch.setProperty "SelectionLanguage", "XPath"
set selectedlistpenSearchnode = doclistpenSearch.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedlistpenSearchnodes=doclistpenSearch.documentElement.selectNodes("/languages/language")
function getlistpenSearchLngStr(instring)
	temp = selectedlistpenSearchnode.selectSingleNode(instring).text
	getlistpenSearchLngStr = temp
end function
%>

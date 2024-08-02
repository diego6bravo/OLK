<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "backToSearch.xml"
set docbackToSearch = server.CreateObject("MSXML2.DOMDocument")
docbackToSearch.async = False
DocbackToSearch.Load(server.MapPath(xmlfilename)) 
docbackToSearch.setProperty "SelectionLanguage", "XPath"
set selectedbackToSearchnode = docbackToSearch.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedbackToSearchnodes=docbackToSearch.documentElement.selectNodes("/languages/language")
function getbackToSearchLngStr(instring)
	temp = selectedbackToSearchnode.selectSingleNode(instring).text
	getbackToSearchLngStr = temp
end function
%>

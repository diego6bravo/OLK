<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "searchClients.xml"
set docsearchClients = server.CreateObject("MSXML2.DOMDocument")
docsearchClients.async = False
DocsearchClients.Load(server.MapPath(xmlfilename)) 
docsearchClients.setProperty "SelectionLanguage", "XPath"
set selectedsearchClientsnode = docsearchClients.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedsearchClientsnodes=docsearchClients.documentElement.selectNodes("/languages/language")
function getsearchClientsLngStr(instring)
	temp = selectedsearchClientsnode.selectSingleNode(instring).text
	getsearchClientsLngStr = temp
end function
%>

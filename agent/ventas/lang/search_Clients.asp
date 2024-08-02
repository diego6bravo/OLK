<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "search_Clients.xml"
set docsearch_Clients = server.CreateObject("MSXML2.DOMDocument")
docsearch_Clients.async = False
Docsearch_Clients.Load(server.MapPath(xmlfilename)) 
docsearch_Clients.setProperty "SelectionLanguage", "XPath"
set selectedsearch_Clientsnode = docsearch_Clients.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedsearch_Clientsnodes=docsearch_Clients.documentElement.selectNodes("/languages/language")
function getsearch_ClientsLngStr(instring)
	temp = selectedsearch_Clientsnode.selectSingleNode(instring).text
	getsearch_ClientsLngStr = temp
end function
%>

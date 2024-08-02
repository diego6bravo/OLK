<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "selectClientes.xml"
set docselectClientes = server.CreateObject("MSXML2.DOMDocument")
docselectClientes.async = False
DocselectClientes.Load(server.MapPath(xmlfilename)) 
docselectClientes.setProperty "SelectionLanguage", "XPath"
set selectedselectClientesnode = docselectClientes.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedselectClientesnodes=docselectClientes.documentElement.selectNodes("/languages/language")
function getselectClientesLngStr(instring)
	temp = selectedselectClientesnode.selectSingleNode(instring).text
	getselectClientesLngStr = temp
end function
%>

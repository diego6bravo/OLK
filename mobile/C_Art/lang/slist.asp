<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "slist.xml"
set docslist = server.CreateObject("MSXML2.DOMDocument")
docslist.async = False
Docslist.Load(server.MapPath(xmlfilename)) 
docslist.setProperty "SelectionLanguage", "XPath"
set selectedslistnode = docslist.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedslistnodes=docslist.documentElement.selectNodes("/languages/language")
function getslistLngStr(instring)
	temp = selectedslistnode.selectSingleNode(instring).text
	getslistLngStr = temp
end function
%>

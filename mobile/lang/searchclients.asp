<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "searchclients.xml"
set docsearchclients = server.CreateObject("MSXML2.DOMDocument")
docsearchclients.async = False
Docsearchclients.Load(server.MapPath(xmlfilename)) 
docsearchclients.setProperty "SelectionLanguage", "XPath"
set selectedsearchclientsnode = docsearchclients.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedsearchclientsnodes=docsearchclients.documentElement.selectNodes("/languages/language")
function getsearchclientsLngStr(instring)
	temp = selectedsearchclientsnode.selectSingleNode(instring).text
	getsearchclientsLngStr = temp
end function
%>

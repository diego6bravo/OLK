<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "searchclientsa.xml"
set docsearchclientsa = server.CreateObject("MSXML2.DOMDocument")
docsearchclientsa.async = False
Docsearchclientsa.Load(server.MapPath(xmlfilename)) 
docsearchclientsa.setProperty "SelectionLanguage", "XPath"
set selectedsearchclientsanode = docsearchclientsa.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedsearchclientsanodes=docsearchclientsa.documentElement.selectNodes("/languages/language")
function getsearchclientsaLngStr(instring)
	temp = selectedsearchclientsanode.selectSingleNode(instring).text
	getsearchclientsaLngStr = temp
end function
%>

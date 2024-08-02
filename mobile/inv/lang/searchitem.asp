<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "searchitem.xml"
set docsearchitem = server.CreateObject("MSXML2.DOMDocument")
docsearchitem.async = False
Docsearchitem.Load(server.MapPath(xmlfilename)) 
docsearchitem.setProperty "SelectionLanguage", "XPath"
set selectedsearchitemnode = docsearchitem.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedsearchitemnodes=docsearchitem.documentElement.selectNodes("/languages/language")
function getsearchitemLngStr(instring)
	temp = selectedsearchitemnode.selectSingleNode(instring).text
	getsearchitemLngStr = temp
end function
%>

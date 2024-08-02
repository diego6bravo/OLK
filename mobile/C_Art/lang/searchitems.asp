<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "searchitems.xml"
set docsearchitems = server.CreateObject("MSXML2.DOMDocument")
docsearchitems.async = False
Docsearchitems.Load(server.MapPath(xmlfilename)) 
docsearchitems.setProperty "SelectionLanguage", "XPath"
set selectedsearchitemsnode = docsearchitems.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedsearchitemsnodes=docsearchitems.documentElement.selectNodes("/languages/language")
function getsearchitemsLngStr(instring)
	temp = selectedsearchitemsnode.selectSingleNode(instring).text
	getsearchitemsLngStr = temp
end function
%>

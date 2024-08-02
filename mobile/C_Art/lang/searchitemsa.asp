<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "searchitemsa.xml"
set docsearchitemsa = server.CreateObject("MSXML2.DOMDocument")
docsearchitemsa.async = False
Docsearchitemsa.Load(server.MapPath(xmlfilename)) 
docsearchitemsa.setProperty "SelectionLanguage", "XPath"
set selectedsearchitemsanode = docsearchitemsa.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedsearchitemsanodes=docsearchitemsa.documentElement.selectNodes("/languages/language")
function getsearchitemsaLngStr(instring)
	temp = selectedsearchitemsanode.selectSingleNode(instring).text
	getsearchitemsaLngStr = temp
end function
%>

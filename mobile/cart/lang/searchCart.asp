<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "searchCart.xml"
set docsearchCart = server.CreateObject("MSXML2.DOMDocument")
docsearchCart.async = False
DocsearchCart.Load(server.MapPath(xmlfilename)) 
docsearchCart.setProperty "SelectionLanguage", "XPath"
set selectedsearchCartnode = docsearchCart.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedsearchCartnodes=docsearchCart.documentElement.selectNodes("/languages/language")
function getsearchCartLngStr(instring)
	temp = selectedsearchCartnode.selectSingleNode(instring).text
	getsearchCartLngStr = temp
end function
%>

<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "payCash.xml"
set docpayCash = server.CreateObject("MSXML2.DOMDocument")
docpayCash.async = False
DocpayCash.Load(server.MapPath(xmlfilename)) 
docpayCash.setProperty "SelectionLanguage", "XPath"
set selectedpayCashnode = docpayCash.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedpayCashnodes=docpayCash.documentElement.selectNodes("/languages/language")
function getpayCashLngStr(instring)
	temp = selectedpayCashnode.selectSingleNode(instring).text
	getpayCashLngStr = temp
end function
%>

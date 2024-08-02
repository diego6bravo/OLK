<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "wish.xml"
set docwish = server.CreateObject("MSXML2.DOMDocument")
docwish.async = False
Docwish.Load(server.MapPath(xmlfilename)) 
docwish.setProperty "SelectionLanguage", "XPath"
set selectedwishnode = docwish.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedwishnodes=docwish.documentElement.selectNodes("/languages/language")
function getwishLngStr(instring)
	temp = selectedwishnode.selectSingleNode(instring).text
	getwishLngStr = temp
end function
%>

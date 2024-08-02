<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "intro.xml"
set docintro = server.CreateObject("MSXML2.DOMDocument")
docintro.async = False
Docintro.Load(server.MapPath(xmlfilename)) 
docintro.setProperty "SelectionLanguage", "XPath"
set selectedintronode = docintro.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedintronodes=docintro.documentElement.selectNodes("/languages/language")
function getintroLngStr(instring)
	temp = selectedintronode.selectSingleNode(instring).text
	getintroLngStr = temp
end function
%>

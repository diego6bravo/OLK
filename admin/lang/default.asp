<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "default.xml"
set docdefault = server.CreateObject("MSXML2.DOMDocument")
docdefault.async = False
Docdefault.Load(server.MapPath(xmlfilename)) 
docdefault.setProperty "SelectionLanguage", "XPath"
set selecteddefaultnode = docdefault.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selecteddefaultnodes=docdefault.documentElement.selectNodes("/languages/language")
function getdefaultLngStr(instring)
	temp = selecteddefaultnode.selectSingleNode(instring).text
	getdefaultLngStr = temp
end function
%>

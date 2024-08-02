<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "accessDenied.xml"
set docaccessDenied = server.CreateObject("MSXML2.DOMDocument")
docaccessDenied.async = False
DocaccessDenied.Load(server.MapPath(xmlfilename)) 
docaccessDenied.setProperty "SelectionLanguage", "XPath"
set selectedaccessDeniednode = docaccessDenied.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedaccessDeniednodes=docaccessDenied.documentElement.selectNodes("/languages/language")
function getaccessDeniedLngStr(instring)
	temp = selectedaccessDeniednode.selectSingleNode(instring).text
	getaccessDeniedLngStr = temp
end function
%>

<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "so.xml"
set docso = server.CreateObject("MSXML2.DOMDocument")
docso.async = False
Docso.Load(server.MapPath(xmlfilename)) 
docso.setProperty "SelectionLanguage", "XPath"
set selectedsonode = docso.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedsonodes=docso.documentElement.selectNodes("/languages/language")
function getsoLngStr(instring)
	temp = selectedsonode.selectSingleNode(instring).text
	getsoLngStr = temp
end function
%>

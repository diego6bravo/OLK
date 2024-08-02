<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "prom.xml"
set docprom = server.CreateObject("MSXML2.DOMDocument")
docprom.async = False
Docprom.Load(server.MapPath(xmlfilename)) 
docprom.setProperty "SelectionLanguage", "XPath"
set selectedpromnode = docprom.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedpromnodes=docprom.documentElement.selectNodes("/languages/language")
function getpromLngStr(instring)
	temp = selectedpromnode.selectSingleNode(instring).text
	getpromLngStr = temp
end function
%>

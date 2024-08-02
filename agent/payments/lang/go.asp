<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "go.xml"
set docgo = server.CreateObject("MSXML2.DOMDocument")
docgo.async = False
Docgo.Load(server.MapPath(xmlfilename)) 
docgo.setProperty "SelectionLanguage", "XPath"
set selectedgonode = docgo.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedgonodes=docgo.documentElement.selectNodes("/languages/language")
function getgoLngStr(instring)
	temp = selectedgonode.selectSingleNode(instring).text
	getgoLngStr = temp
end function
%>

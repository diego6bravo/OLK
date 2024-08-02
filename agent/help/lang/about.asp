<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "about.xml"
set docabout = server.CreateObject("MSXML2.DOMDocument")
docabout.async = False
Docabout.Load(server.MapPath(xmlfilename)) 
docabout.setProperty "SelectionLanguage", "XPath"
set selectedaboutnode = docabout.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedaboutnodes=docabout.documentElement.selectNodes("/languages/language")
function getaboutLngStr(instring)
	temp = selectedaboutnode.selectSingleNode(instring).text
	getaboutLngStr = temp
end function
%>

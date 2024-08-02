<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "buzon.xml"
set docbuzon = server.CreateObject("MSXML2.DOMDocument")
docbuzon.async = False
Docbuzon.Load(server.MapPath(xmlfilename)) 
docbuzon.setProperty "SelectionLanguage", "XPath"
set selectedbuzonnode = docbuzon.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedbuzonnodes=docbuzon.documentElement.selectNodes("/languages/language")
function getbuzonLngStr(instring)
	temp = selectedbuzonnode.selectSingleNode(instring).text
	getbuzonLngStr = temp
end function
%>

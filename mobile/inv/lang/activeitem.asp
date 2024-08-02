<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "activeitem.xml"
set docactiveitem = server.CreateObject("MSXML2.DOMDocument")
docactiveitem.async = False
Docactiveitem.Load(server.MapPath(xmlfilename)) 
docactiveitem.setProperty "SelectionLanguage", "XPath"
set selectedactiveitemnode = docactiveitem.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedactiveitemnodes=docactiveitem.documentElement.selectNodes("/languages/language")
function getactiveitemLngStr(instring)
	temp = selectedactiveitemnode.selectSingleNode(instring).text
	getactiveitemLngStr = temp
end function
%>

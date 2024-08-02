<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "maps.xml"
set docmaps = server.CreateObject("MSXML2.DOMDocument")
docmaps.async = False
Docmaps.Load(server.MapPath(xmlfilename)) 
docmaps.setProperty "SelectionLanguage", "XPath"
set selectedmapsnode = docmaps.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedmapsnodes=docmaps.documentElement.selectNodes("/languages/language")
function getmapsLngStr(instring)
	temp = selectedmapsnode.selectSingleNode(instring).text
	getmapsLngStr = temp
end function
%>

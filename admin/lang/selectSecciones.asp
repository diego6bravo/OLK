<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "selectSecciones.xml"
set docselectSecciones = server.CreateObject("MSXML2.DOMDocument")
docselectSecciones.async = False
DocselectSecciones.Load(server.MapPath(xmlfilename)) 
docselectSecciones.setProperty "SelectionLanguage", "XPath"
set selectedselectSeccionesnode = docselectSecciones.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedselectSeccionesnodes=docselectSecciones.documentElement.selectNodes("/languages/language")
function getselectSeccionesLngStr(instring)
	temp = selectedselectSeccionesnode.selectSingleNode(instring).text
	getselectSeccionesLngStr = temp
end function
%>

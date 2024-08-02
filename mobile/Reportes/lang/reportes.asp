<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "reportes.xml"
set docreportes = server.CreateObject("MSXML2.DOMDocument")
docreportes.async = False
Docreportes.Load(server.MapPath(xmlfilename)) 
docreportes.setProperty "SelectionLanguage", "XPath"
set selectedreportesnode = docreportes.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedreportesnodes=docreportes.documentElement.selectNodes("/languages/language")
function getreportesLngStr(instring)
	temp = selectedreportesnode.selectSingleNode(instring).text
	getreportesLngStr = temp
end function
%>

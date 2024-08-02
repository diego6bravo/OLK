<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminRepExport.xml"
set docadminRepExport = server.CreateObject("MSXML2.DOMDocument")
docadminRepExport.async = False
DocadminRepExport.Load(server.MapPath(xmlfilename)) 
docadminRepExport.setProperty "SelectionLanguage", "XPath"
set selectedadminRepExportnode = docadminRepExport.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminRepExportnodes=docadminRepExport.documentElement.selectNodes("/languages/language")
function getadminRepExportLngStr(instring)
	temp = selectedadminRepExportnode.selectSingleNode(instring).text
	getadminRepExportLngStr = temp
end function
%>

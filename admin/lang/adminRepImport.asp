<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminRepImport.xml"
set docadminRepImport = server.CreateObject("MSXML2.DOMDocument")
docadminRepImport.async = False
DocadminRepImport.Load(server.MapPath(xmlfilename)) 
docadminRepImport.setProperty "SelectionLanguage", "XPath"
set selectedadminRepImportnode = docadminRepImport.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminRepImportnodes=docadminRepImport.documentElement.selectNodes("/languages/language")
function getadminRepImportLngStr(instring)
	temp = selectedadminRepImportnode.selectSingleNode(instring).text
	getadminRepImportLngStr = temp
end function
%>

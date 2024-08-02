<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "cartImport.xml"
set doccartImport = server.CreateObject("MSXML2.DOMDocument")
doccartImport.async = False
DoccartImport.Load(server.MapPath(xmlfilename)) 
doccartImport.setProperty "SelectionLanguage", "XPath"
set selectedcartImportnode = doccartImport.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedcartImportnodes=doccartImport.documentElement.selectNodes("/languages/language")
function getcartImportLngStr(instring)
	temp = selectedcartImportnode.selectSingleNode(instring).text
	getcartImportLngStr = temp
end function
%>

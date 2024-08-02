<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminObjPrint.xml"
set docadminObjPrint = server.CreateObject("MSXML2.DOMDocument")
docadminObjPrint.async = False
DocadminObjPrint.Load(server.MapPath(xmlfilename)) 
docadminObjPrint.setProperty "SelectionLanguage", "XPath"
set selectedadminObjPrintnode = docadminObjPrint.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminObjPrintnodes=docadminObjPrint.documentElement.selectNodes("/languages/language")
function getadminObjPrintLngStr(instring)
	temp = selectedadminObjPrintnode.selectSingleNode(instring).text
	getadminObjPrintLngStr = temp
end function
%>

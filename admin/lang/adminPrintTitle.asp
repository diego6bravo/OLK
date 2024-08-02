<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminPrintTitle.xml"
set docadminPrintTitle = server.CreateObject("MSXML2.DOMDocument")
docadminPrintTitle.async = False
DocadminPrintTitle.Load(server.MapPath(xmlfilename)) 
docadminPrintTitle.setProperty "SelectionLanguage", "XPath"
set selectedadminPrintTitlenode = docadminPrintTitle.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminPrintTitlenodes=docadminPrintTitle.documentElement.selectNodes("/languages/language")
function getadminPrintTitleLngStr(instring)
	temp = selectedadminPrintTitlenode.selectSingleNode(instring).text
	getadminPrintTitleLngStr = temp
end function
%>

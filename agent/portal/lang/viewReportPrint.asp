<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "viewReportPrint.xml"
set docviewReportPrint = server.CreateObject("MSXML2.DOMDocument")
docviewReportPrint.async = False
DocviewReportPrint.Load(server.MapPath(xmlfilename)) 
docviewReportPrint.setProperty "SelectionLanguage", "XPath"
set selectedviewReportPrintnode = docviewReportPrint.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedviewReportPrintnodes=docviewReportPrint.documentElement.selectNodes("/languages/language")
function getviewReportPrintLngStr(instring)
	temp = selectedviewReportPrintnode.selectSingleNode(instring).text
	getviewReportPrintLngStr = temp
end function
%>

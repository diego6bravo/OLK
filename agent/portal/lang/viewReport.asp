<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "viewReport.xml"
set docviewReport = server.CreateObject("MSXML2.DOMDocument")
docviewReport.async = False
DocviewReport.Load(server.MapPath(xmlfilename)) 
docviewReport.setProperty "SelectionLanguage", "XPath"
set selectedviewReportnode = docviewReport.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedviewReportnodes=docviewReport.documentElement.selectNodes("/languages/language")
function getviewReportLngStr(instring)
	temp = selectedviewReportnode.selectSingleNode(instring).text
	getviewReportLngStr = temp
end function
%>

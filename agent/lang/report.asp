<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "report.xml"
set docreport = server.CreateObject("MSXML2.DOMDocument")
docreport.async = False
Docreport.Load(server.MapPath(xmlfilename)) 
docreport.setProperty "SelectionLanguage", "XPath"
set selectedreportnode = docreport.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedreportnodes=docreport.documentElement.selectNodes("/languages/language")
function getreportLngStr(instring)
	temp = selectedreportnode.selectSingleNode(instring).text
	getreportLngStr = temp
end function
%>

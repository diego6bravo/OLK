<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminOps.xml"
set docadminOps = server.CreateObject("MSXML2.DOMDocument")
docadminOps.async = False
DocadminOps.Load(server.MapPath(xmlfilename)) 
docadminOps.setProperty "SelectionLanguage", "XPath"
set selectedadminOpsnode = docadminOps.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminOpsnodes=docadminOps.documentElement.selectNodes("/languages/language")
function getadminOpsLngStr(instring)
	temp = selectedadminOpsnode.selectSingleNode(instring).text
	getadminOpsLngStr = temp
end function
%>

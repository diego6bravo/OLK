<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminCardOpt.xml"
set docadminCardOpt = server.CreateObject("MSXML2.DOMDocument")
docadminCardOpt.async = False
DocadminCardOpt.Load(server.MapPath(xmlfilename)) 
docadminCardOpt.setProperty "SelectionLanguage", "XPath"
set selectedadminCardOptnode = docadminCardOpt.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminCardOptnodes=docadminCardOpt.documentElement.selectNodes("/languages/language")
function getadminCardOptLngStr(instring)
	temp = selectedadminCardOptnode.selectSingleNode(instring).text
	getadminCardOptLngStr = temp
end function
%>

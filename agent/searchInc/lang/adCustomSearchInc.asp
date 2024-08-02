<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adCustomSearchInc.xml"
set docadCustomSearchInc = server.CreateObject("MSXML2.DOMDocument")
docadCustomSearchInc.async = False
DocadCustomSearchInc.Load(server.MapPath(xmlfilename)) 
docadCustomSearchInc.setProperty "SelectionLanguage", "XPath"
set selectedadCustomSearchIncnode = docadCustomSearchInc.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadCustomSearchIncnodes=docadCustomSearchInc.documentElement.selectNodes("/languages/language")
function getadCustomSearchIncLngStr(instring)
	temp = selectedadCustomSearchIncnode.selectSingleNode(instring).text
	getadCustomSearchIncLngStr = temp
end function
%>

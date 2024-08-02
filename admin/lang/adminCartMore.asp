<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminCartMore.xml"
set docadminCartMore = server.CreateObject("MSXML2.DOMDocument")
docadminCartMore.async = False
DocadminCartMore.Load(server.MapPath(xmlfilename)) 
docadminCartMore.setProperty "SelectionLanguage", "XPath"
set selectedadminCartMorenode = docadminCartMore.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminCartMorenodes=docadminCartMore.documentElement.selectNodes("/languages/language")
function getadminCartMoreLngStr(instring)
	temp = selectedadminCartMorenode.selectSingleNode(instring).text
	getadminCartMoreLngStr = temp
end function
%>

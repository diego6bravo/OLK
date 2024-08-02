<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminUpdate.xml"
set docadminUpdate = server.CreateObject("MSXML2.DOMDocument")
docadminUpdate.async = False
DocadminUpdate.Load(server.MapPath(xmlfilename)) 
docadminUpdate.setProperty "SelectionLanguage", "XPath"
set selectedadminUpdatenode = docadminUpdate.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminUpdatenodes=docadminUpdate.documentElement.selectNodes("/languages/language")
function getadminUpdateLngStr(instring)
	temp = selectedadminUpdatenode.selectSingleNode(instring).text
	getadminUpdateLngStr = temp
end function
%>

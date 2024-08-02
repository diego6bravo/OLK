<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminPriceCod.xml"
set docadminPriceCod = server.CreateObject("MSXML2.DOMDocument")
docadminPriceCod.async = False
DocadminPriceCod.Load(server.MapPath(xmlfilename)) 
docadminPriceCod.setProperty "SelectionLanguage", "XPath"
set selectedadminPriceCodnode = docadminPriceCod.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminPriceCodnodes=docadminPriceCod.documentElement.selectNodes("/languages/language")
function getadminPriceCodLngStr(instring)
	temp = selectedadminPriceCodnode.selectSingleNode(instring).text
	getadminPriceCodLngStr = temp
end function
%>

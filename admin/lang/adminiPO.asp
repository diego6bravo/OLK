<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminiPO.xml"
set docadminiPO = server.CreateObject("MSXML2.DOMDocument")
docadminiPO.async = False
DocadminiPO.Load(server.MapPath(xmlfilename)) 
docadminiPO.setProperty "SelectionLanguage", "XPath"
set selectedadminiPOnode = docadminiPO.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminiPOnodes=docadminiPO.documentElement.selectNodes("/languages/language")
function getadminiPOLngStr(instring)
	temp = selectedadminiPOnode.selectSingleNode(instring).text
	getadminiPOLngStr = temp
end function
%>

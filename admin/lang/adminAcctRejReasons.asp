<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminAcctRejReasons.xml"
set docadminAcctRejReasons = server.CreateObject("MSXML2.DOMDocument")
docadminAcctRejReasons.async = False
DocadminAcctRejReasons.Load(server.MapPath(xmlfilename)) 
docadminAcctRejReasons.setProperty "SelectionLanguage", "XPath"
set selectedadminAcctRejReasonsnode = docadminAcctRejReasons.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminAcctRejReasonsnodes=docadminAcctRejReasons.documentElement.selectNodes("/languages/language")
function getadminAcctRejReasonsLngStr(instring)
	temp = selectedadminAcctRejReasonsnode.selectSingleNode(instring).text
	getadminAcctRejReasonsLngStr = temp
end function
%>

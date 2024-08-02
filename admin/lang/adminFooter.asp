<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminFooter.xml"
set docadminFooter = server.CreateObject("MSXML2.DOMDocument")
docadminFooter.async = False
DocadminFooter.Load(server.MapPath(xmlfilename)) 
docadminFooter.setProperty "SelectionLanguage", "XPath"
set selectedadminFooternode = docadminFooter.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminFooternodes=docadminFooter.documentElement.selectNodes("/languages/language")
function getadminFooterLngStr(instring)
	temp = selectedadminFooternode.selectSingleNode(instring).text
	getadminFooterLngStr = temp
end function
%>

<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminPaySis.xml"
set docadminPaySis = server.CreateObject("MSXML2.DOMDocument")
docadminPaySis.async = False
DocadminPaySis.Load(server.MapPath(xmlfilename)) 
docadminPaySis.setProperty "SelectionLanguage", "XPath"
set selectedadminPaySisnode = docadminPaySis.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminPaySisnodes=docadminPaySis.documentElement.selectNodes("/languages/language")
function getadminPaySisLngStr(instring)
	temp = selectedadminPaySisnode.selectSingleNode(instring).text
	getadminPaySisLngStr = temp
end function
%>

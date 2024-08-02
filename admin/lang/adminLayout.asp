<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminLayout.xml"
set docadminLayout = server.CreateObject("MSXML2.DOMDocument")
docadminLayout.async = False
DocadminLayout.Load(server.MapPath(xmlfilename)) 
docadminLayout.setProperty "SelectionLanguage", "XPath"
set selectedadminLayoutnode = docadminLayout.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminLayoutnodes=docadminLayout.documentElement.selectNodes("/languages/language")
function getadminLayoutLngStr(instring)
	temp = selectedadminLayoutnode.selectSingleNode(instring).text
	getadminLayoutLngStr = temp
end function
%>

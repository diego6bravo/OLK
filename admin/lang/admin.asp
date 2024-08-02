<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "admin.xml"
set docadmin = server.CreateObject("MSXML2.DOMDocument")
docadmin.async = False
Docadmin.Load(server.MapPath(xmlfilename)) 
docadmin.setProperty "SelectionLanguage", "XPath"
set selectedadminnode = docadmin.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminnodes=docadmin.documentElement.selectNodes("/languages/language")
function getadminLngStr(instring)
	temp = selectedadminnode.selectSingleNode(instring).text
	getadminLngStr = temp
end function
%>

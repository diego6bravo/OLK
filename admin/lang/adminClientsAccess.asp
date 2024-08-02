<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminClientsAccess.xml"
set docadminClientsAccess = server.CreateObject("MSXML2.DOMDocument")
docadminClientsAccess.async = False
DocadminClientsAccess.Load(server.MapPath(xmlfilename)) 
docadminClientsAccess.setProperty "SelectionLanguage", "XPath"
set selectedadminClientsAccessnode = docadminClientsAccess.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminClientsAccessnodes=docadminClientsAccess.documentElement.selectNodes("/languages/language")
function getadminClientsAccessLngStr(instring)
	temp = selectedadminClientsAccessnode.selectSingleNode(instring).text
	getadminClientsAccessLngStr = temp
end function
%>

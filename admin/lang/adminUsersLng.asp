<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminUsersLng.xml"
set docadminUsersLng = server.CreateObject("MSXML2.DOMDocument")
docadminUsersLng.async = False
DocadminUsersLng.Load(server.MapPath(xmlfilename)) 
docadminUsersLng.setProperty "SelectionLanguage", "XPath"
set selectedadminUsersLngnode = docadminUsersLng.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminUsersLngnodes=docadminUsersLng.documentElement.selectNodes("/languages/language")
function getadminUsersLngLngStr(instring)
	temp = selectedadminUsersLngnode.selectSingleNode(instring).text
	getadminUsersLngLngStr = temp
end function
%>

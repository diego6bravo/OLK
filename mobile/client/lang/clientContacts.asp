<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "clientContacts.xml"
set docclientContacts = server.CreateObject("MSXML2.DOMDocument")
docclientContacts.async = False
DocclientContacts.Load(server.MapPath(xmlfilename)) 
docclientContacts.setProperty "SelectionLanguage", "XPath"
set selectedclientContactsnode = docclientContacts.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedclientContactsnodes=docclientContacts.documentElement.selectNodes("/languages/language")
function getclientContactsLngStr(instring)
	temp = selectedclientContactsnode.selectSingleNode(instring).text
	getclientContactsLngStr = temp
end function
%>

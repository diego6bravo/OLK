<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "contacts.xml"
set doccontacts = server.CreateObject("MSXML2.DOMDocument")
doccontacts.async = False
Doccontacts.Load(server.MapPath(xmlfilename)) 
doccontacts.setProperty "SelectionLanguage", "XPath"
set selectedcontactsnode = doccontacts.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedcontactsnodes=doccontacts.documentElement.selectNodes("/languages/language")
function getcontactsLngStr(instring)
	temp = selectedcontactsnode.selectSingleNode(instring).text
	getcontactsLngStr = temp
end function
%>

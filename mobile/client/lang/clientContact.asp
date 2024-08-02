<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "clientContact.xml"
set docclientContact = server.CreateObject("MSXML2.DOMDocument")
docclientContact.async = False
DocclientContact.Load(server.MapPath(xmlfilename)) 
docclientContact.setProperty "SelectionLanguage", "XPath"
set selectedclientContactnode = docclientContact.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedclientContactnodes=docclientContact.documentElement.selectNodes("/languages/language")
function getclientContactLngStr(instring)
	temp = selectedclientContactnode.selectSingleNode(instring).text
	getclientContactLngStr = temp
end function
%>

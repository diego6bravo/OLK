<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "sapUsers.xml"
set docsapUsers = server.CreateObject("MSXML2.DOMDocument")
docsapUsers.async = False
DocsapUsers.Load(server.MapPath(xmlfilename)) 
docsapUsers.setProperty "SelectionLanguage", "XPath"
set selectedsapUsersnode = docsapUsers.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedsapUsersnodes=docsapUsers.documentElement.selectNodes("/languages/language")
function getsapUsersLngStr(instring)
	temp = selectedsapUsersnode.selectSingleNode(instring).text
	getsapUsersLngStr = temp
end function
%>

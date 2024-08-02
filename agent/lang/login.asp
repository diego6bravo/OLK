<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "login.xml"
set doclogin = server.CreateObject("MSXML2.DOMDocument")
doclogin.async = False
Doclogin.Load(server.MapPath(xmlfilename)) 
doclogin.setProperty "SelectionLanguage", "XPath"
set selectedloginnode = doclogin.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedloginnodes=doclogin.documentElement.selectNodes("/languages/language")
function getloginLngStr(instring)
	temp = selectedloginnode.selectSingleNode(instring).text
	getloginLngStr = temp
end function
%>

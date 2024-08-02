<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "errNoAccess.xml"
set docerrNoAccess = server.CreateObject("MSXML2.DOMDocument")
docerrNoAccess.async = False
DocerrNoAccess.Load(server.MapPath(xmlfilename)) 
docerrNoAccess.setProperty "SelectionLanguage", "XPath"
set selectederrNoAccessnode = docerrNoAccess.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectederrNoAccessnodes=docerrNoAccess.documentElement.selectNodes("/languages/language")
function geterrNoAccessLngStr(instring)
	temp = selectederrNoAccessnode.selectSingleNode(instring).text
	geterrNoAccessLngStr = temp
end function
%>

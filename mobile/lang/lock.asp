<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "lock.xml"
set doclock = server.CreateObject("MSXML2.DOMDocument")
doclock.async = False
Doclock.Load(server.MapPath(xmlfilename)) 
doclock.setProperty "SelectionLanguage", "XPath"
set selectedlocknode = doclock.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedlocknodes=doclock.documentElement.selectNodes("/languages/language")
function getlockLngStr(instring)
	temp = selectedlocknode.selectSingleNode(instring).text
	getlockLngStr = temp
end function
%>

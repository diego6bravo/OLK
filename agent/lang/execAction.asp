<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "execAction.xml"
set docexecAction = server.CreateObject("MSXML2.DOMDocument")
docexecAction.async = False
DocexecAction.Load(server.MapPath(xmlfilename)) 
docexecAction.setProperty "SelectionLanguage", "XPath"
set selectedexecActionnode = docexecAction.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedexecActionnodes=docexecAction.documentElement.selectNodes("/languages/language")
function getexecActionLngStr(instring)
	temp = selectedexecActionnode.selectSingleNode(instring).text
	getexecActionLngStr = temp
end function
%>

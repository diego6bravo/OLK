<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "passwordReset.xml"
set docpasswordReset = server.CreateObject("MSXML2.DOMDocument")
docpasswordReset.async = False
DocpasswordReset.Load(server.MapPath(xmlfilename)) 
docpasswordReset.setProperty "SelectionLanguage", "XPath"
set selectedpasswordResetnode = docpasswordReset.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedpasswordResetnodes=docpasswordReset.documentElement.selectNodes("/languages/language")
function getpasswordResetLngStr(instring)
	temp = selectedpasswordResetnode.selectSingleNode(instring).text
	getpasswordResetLngStr = temp
end function
%>

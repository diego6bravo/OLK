<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "autSubmit.xml"
set docautSubmit = server.CreateObject("MSXML2.DOMDocument")
docautSubmit.async = False
DocautSubmit.Load(server.MapPath(xmlfilename)) 
docautSubmit.setProperty "SelectionLanguage", "XPath"
set selectedautSubmitnode = docautSubmit.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedautSubmitnodes=docautSubmit.documentElement.selectNodes("/languages/language")
function getautSubmitLngStr(instring)
	temp = selectedautSubmitnode.selectSingleNode(instring).text
	getautSubmitLngStr = temp
end function
%>

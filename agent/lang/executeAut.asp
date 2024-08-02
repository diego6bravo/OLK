<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "executeAut.xml"
set docexecuteAut = server.CreateObject("MSXML2.DOMDocument")
docexecuteAut.async = False
DocexecuteAut.Load(server.MapPath(xmlfilename)) 
docexecuteAut.setProperty "SelectionLanguage", "XPath"
set selectedexecuteAutnode = docexecuteAut.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedexecuteAutnodes=docexecuteAut.documentElement.selectNodes("/languages/language")
function getexecuteAutLngStr(instring)
	temp = selectedexecuteAutnode.selectSingleNode(instring).text
	getexecuteAutLngStr = temp
end function
%>

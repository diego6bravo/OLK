<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "myFunctions.xml"
set docmyFunctions = server.CreateObject("MSXML2.DOMDocument")
docmyFunctions.async = False
DocmyFunctions.Load(server.MapPath(xmlfilename)) 
docmyFunctions.setProperty "SelectionLanguage", "XPath"
set selectedmyFunctionsnode = docmyFunctions.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedmyFunctionsnodes=docmyFunctions.documentElement.selectNodes("/languages/language")
function getmyFunctionsLngStr(instring)
	temp = selectedmyFunctionsnode.selectSingleNode(instring).text
	getmyFunctionsLngStr = temp
end function
%>

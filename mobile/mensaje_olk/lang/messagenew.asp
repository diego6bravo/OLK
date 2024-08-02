<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "messagenew.xml"
set docmessagenew = server.CreateObject("MSXML2.DOMDocument")
docmessagenew.async = False
Docmessagenew.Load(server.MapPath(xmlfilename)) 
docmessagenew.setProperty "SelectionLanguage", "XPath"
set selectedmessagenewnode = docmessagenew.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedmessagenewnodes=docmessagenew.documentElement.selectNodes("/languages/language")
function getmessagenewLngStr(instring)
	temp = selectedmessagenewnode.selectSingleNode(instring).text
	getmessagenewLngStr = temp
end function
%>

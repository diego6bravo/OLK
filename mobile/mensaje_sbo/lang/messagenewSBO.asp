<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "messagenewSBO.xml"
set docmessagenewSBO = server.CreateObject("MSXML2.DOMDocument")
docmessagenewSBO.async = False
DocmessagenewSBO.Load(server.MapPath(xmlfilename)) 
docmessagenewSBO.setProperty "SelectionLanguage", "XPath"
set selectedmessagenewSBOnode = docmessagenewSBO.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedmessagenewSBOnodes=docmessagenewSBO.documentElement.selectNodes("/languages/language")
function getmessagenewSBOLngStr(instring)
	temp = selectedmessagenewSBOnode.selectSingleNode(instring).text
	getmessagenewSBOLngStr = temp
end function
%>

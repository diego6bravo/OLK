<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "messages.xml"
set docmessages = server.CreateObject("MSXML2.DOMDocument")
docmessages.async = False
Docmessages.Load(server.MapPath(xmlfilename)) 
docmessages.setProperty "SelectionLanguage", "XPath"
set selectedmessagesnode = docmessages.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedmessagesnodes=docmessages.documentElement.selectNodes("/languages/language")
function getmessagesLngStr(instring)
	temp = selectedmessagesnode.selectSingleNode(instring).text
	getmessagesLngStr = temp
end function
%>

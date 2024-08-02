<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "messageNew.xml"
set docmessageNew = server.CreateObject("MSXML2.DOMDocument")
docmessageNew.async = False
DocmessageNew.Load(server.MapPath(xmlfilename)) 
docmessageNew.setProperty "SelectionLanguage", "XPath"
set selectedmessageNewnode = docmessageNew.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedmessageNewnodes=docmessageNew.documentElement.selectNodes("/languages/language")
function getmessageNewLngStr(instring)
	temp = selectedmessageNewnode.selectSingleNode(instring).text
	getmessageNewLngStr = temp
end function
%>

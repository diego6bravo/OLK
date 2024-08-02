<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "messageReturn.xml"
set docmessageReturn = server.CreateObject("MSXML2.DOMDocument")
docmessageReturn.async = False
DocmessageReturn.Load(server.MapPath(xmlfilename)) 
docmessageReturn.setProperty "SelectionLanguage", "XPath"
set selectedmessageReturnnode = docmessageReturn.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedmessageReturnnodes=docmessageReturn.documentElement.selectNodes("/languages/language")
function getmessageReturnLngStr(instring)
	temp = selectedmessageReturnnode.selectSingleNode(instring).text
	getmessageReturnLngStr = temp
end function
%>

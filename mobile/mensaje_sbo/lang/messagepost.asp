<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "messagepost.xml"
set docmessagepost = server.CreateObject("MSXML2.DOMDocument")
docmessagepost.async = False
Docmessagepost.Load(server.MapPath(xmlfilename)) 
docmessagepost.setProperty "SelectionLanguage", "XPath"
set selectedmessagepostnode = docmessagepost.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedmessagepostnodes=docmessagepost.documentElement.selectNodes("/languages/language")
function getmessagepostLngStr(instring)
	temp = selectedmessagepostnode.selectSingleNode(instring).text
	getmessagepostLngStr = temp
end function
%>

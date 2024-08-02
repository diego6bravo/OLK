<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "messagePost.xml"
set docmessagePost = server.CreateObject("MSXML2.DOMDocument")
docmessagePost.async = False
DocmessagePost.Load(server.MapPath(xmlfilename)) 
docmessagePost.setProperty "SelectionLanguage", "XPath"
set selectedmessagePostnode = docmessagePost.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedmessagePostnodes=docmessagePost.documentElement.selectNodes("/languages/language")
function getmessagePostLngStr(instring)
	temp = selectedmessagePostnode.selectSingleNode(instring).text
	getmessagePostLngStr = temp
end function
%>

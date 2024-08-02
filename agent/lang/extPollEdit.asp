<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "extPollEdit.xml"
set docextPollEdit = server.CreateObject("MSXML2.DOMDocument")
docextPollEdit.async = False
DocextPollEdit.Load(server.MapPath(xmlfilename)) 
docextPollEdit.setProperty "SelectionLanguage", "XPath"
set selectedextPollEditnode = docextPollEdit.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedextPollEditnodes=docextPollEdit.documentElement.selectNodes("/languages/language")
function getextPollEditLngStr(instring)
	temp = selectedextPollEditnode.selectSingleNode(instring).text
	getextPollEditLngStr = temp
end function
%>

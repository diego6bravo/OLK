<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminPollEdit.xml"
set docadminPollEdit = server.CreateObject("MSXML2.DOMDocument")
docadminPollEdit.async = False
DocadminPollEdit.Load(server.MapPath(xmlfilename)) 
docadminPollEdit.setProperty "SelectionLanguage", "XPath"
set selectedadminPollEditnode = docadminPollEdit.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminPollEditnodes=docadminPollEdit.documentElement.selectNodes("/languages/language")
function getadminPollEditLngStr(instring)
	temp = selectedadminPollEditnode.selectSingleNode(instring).text
	getadminPollEditLngStr = temp
end function
%>

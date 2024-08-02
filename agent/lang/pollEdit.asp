<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "pollEdit.xml"
set docpollEdit = server.CreateObject("MSXML2.DOMDocument")
docpollEdit.async = False
DocpollEdit.Load(server.MapPath(xmlfilename)) 
docpollEdit.setProperty "SelectionLanguage", "XPath"
set selectedpollEditnode = docpollEdit.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedpollEditnodes=docpollEdit.documentElement.selectNodes("/languages/language")
function getpollEditLngStr(instring)
	temp = selectedpollEditnode.selectSingleNode(instring).text
	getpollEditLngStr = temp
end function
%>

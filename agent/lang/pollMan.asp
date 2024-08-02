<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "pollMan.xml"
set docpollMan = server.CreateObject("MSXML2.DOMDocument")
docpollMan.async = False
DocpollMan.Load(server.MapPath(xmlfilename)) 
docpollMan.setProperty "SelectionLanguage", "XPath"
set selectedpollMannode = docpollMan.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedpollMannodes=docpollMan.documentElement.selectNodes("/languages/language")
function getpollManLngStr(instring)
	temp = selectedpollMannode.selectSingleNode(instring).text
	getpollManLngStr = temp
end function
%>

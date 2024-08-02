<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "extPollMan.xml"
set docextPollMan = server.CreateObject("MSXML2.DOMDocument")
docextPollMan.async = False
DocextPollMan.Load(server.MapPath(xmlfilename)) 
docextPollMan.setProperty "SelectionLanguage", "XPath"
set selectedextPollMannode = docextPollMan.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedextPollMannodes=docextPollMan.documentElement.selectNodes("/languages/language")
function getextPollManLngStr(instring)
	temp = selectedextPollMannode.selectSingleNode(instring).text
	getextPollManLngStr = temp
end function
%>

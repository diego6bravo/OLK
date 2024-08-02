<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "extPollExec.xml"
set docextPollExec = server.CreateObject("MSXML2.DOMDocument")
docextPollExec.async = False
DocextPollExec.Load(server.MapPath(xmlfilename)) 
docextPollExec.setProperty "SelectionLanguage", "XPath"
set selectedextPollExecnode = docextPollExec.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedextPollExecnodes=docextPollExec.documentElement.selectNodes("/languages/language")
function getextPollExecLngStr(instring)
	temp = selectedextPollExecnode.selectSingleNode(instring).text
	getextPollExecLngStr = temp
end function
%>

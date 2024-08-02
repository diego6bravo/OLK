<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "extPollView.xml"
set docextPollView = server.CreateObject("MSXML2.DOMDocument")
docextPollView.async = False
DocextPollView.Load(server.MapPath(xmlfilename)) 
docextPollView.setProperty "SelectionLanguage", "XPath"
set selectedextPollViewnode = docextPollView.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedextPollViewnodes=docextPollView.documentElement.selectNodes("/languages/language")
function getextPollViewLngStr(instring)
	temp = selectedextPollViewnode.selectSingleNode(instring).text
	getextPollViewLngStr = temp
end function
%>

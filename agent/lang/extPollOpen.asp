<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "extPollOpen.xml"
set docextPollOpen = server.CreateObject("MSXML2.DOMDocument")
docextPollOpen.async = False
DocextPollOpen.Load(server.MapPath(xmlfilename)) 
docextPollOpen.setProperty "SelectionLanguage", "XPath"
set selectedextPollOpennode = docextPollOpen.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedextPollOpennodes=docextPollOpen.documentElement.selectNodes("/languages/language")
function getextPollOpenLngStr(instring)
	temp = selectedextPollOpennode.selectSingleNode(instring).text
	getextPollOpenLngStr = temp
end function
%>

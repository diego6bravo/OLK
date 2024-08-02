<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "extPollList.xml"
set docextPollList = server.CreateObject("MSXML2.DOMDocument")
docextPollList.async = False
DocextPollList.Load(server.MapPath(xmlfilename)) 
docextPollList.setProperty "SelectionLanguage", "XPath"
set selectedextPollListnode = docextPollList.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedextPollListnodes=docextPollList.documentElement.selectNodes("/languages/language")
function getextPollListLngStr(instring)
	temp = selectedextPollListnode.selectSingleNode(instring).text
	getextPollListLngStr = temp
end function
%>

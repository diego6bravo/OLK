<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "extPollViewDetails.xml"
set docextPollViewDetails = server.CreateObject("MSXML2.DOMDocument")
docextPollViewDetails.async = False
DocextPollViewDetails.Load(server.MapPath(xmlfilename)) 
docextPollViewDetails.setProperty "SelectionLanguage", "XPath"
set selectedextPollViewDetailsnode = docextPollViewDetails.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedextPollViewDetailsnodes=docextPollViewDetails.documentElement.selectNodes("/languages/language")
function getextPollViewDetailsLngStr(instring)
	temp = selectedextPollViewDetailsnode.selectSingleNode(instring).text
	getextPollViewDetailsLngStr = temp
end function
%>

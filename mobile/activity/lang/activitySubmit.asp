<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "activitySubmit.xml"
set docactivitySubmit = server.CreateObject("MSXML2.DOMDocument")
docactivitySubmit.async = False
DocactivitySubmit.Load(server.MapPath(xmlfilename)) 
docactivitySubmit.setProperty "SelectionLanguage", "XPath"
set selectedactivitySubmitnode = docactivitySubmit.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedactivitySubmitnodes=docactivitySubmit.documentElement.selectNodes("/languages/language")
function getactivitySubmitLngStr(instring)
	temp = selectedactivitySubmitnode.selectSingleNode(instring).text
	getactivitySubmitLngStr = temp
end function
%>

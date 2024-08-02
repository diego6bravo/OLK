<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "activityCancel.xml"
set docactivityCancel = server.CreateObject("MSXML2.DOMDocument")
docactivityCancel.async = False
DocactivityCancel.Load(server.MapPath(xmlfilename)) 
docactivityCancel.setProperty "SelectionLanguage", "XPath"
set selectedactivityCancelnode = docactivityCancel.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedactivityCancelnodes=docactivityCancel.documentElement.selectNodes("/languages/language")
function getactivityCancelLngStr(instring)
	temp = selectedactivityCancelnode.selectSingleNode(instring).text
	getactivityCancelLngStr = temp
end function
%>

<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "activityX.xml"
set docactivityX = server.CreateObject("MSXML2.DOMDocument")
docactivityX.async = False
DocactivityX.Load(server.MapPath(xmlfilename)) 
docactivityX.setProperty "SelectionLanguage", "XPath"
set selectedactivityXnode = docactivityX.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedactivityXnodes=docactivityX.documentElement.selectNodes("/languages/language")
function getactivityXLngStr(instring)
	temp = selectedactivityXnode.selectSingleNode(instring).text
	getactivityXLngStr = temp
end function
%>

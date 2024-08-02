<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "activityContent.xml"
set docactivityContent = server.CreateObject("MSXML2.DOMDocument")
docactivityContent.async = False
DocactivityContent.Load(server.MapPath(xmlfilename)) 
docactivityContent.setProperty "SelectionLanguage", "XPath"
set selectedactivityContentnode = docactivityContent.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedactivityContentnodes=docactivityContent.documentElement.selectNodes("/languages/language")
function getactivityContentLngStr(instring)
	temp = selectedactivityContentnode.selectSingleNode(instring).text
	getactivityContentLngStr = temp
end function
%>

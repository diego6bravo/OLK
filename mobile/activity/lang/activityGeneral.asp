<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "activityGeneral.xml"
set docactivityGeneral = server.CreateObject("MSXML2.DOMDocument")
docactivityGeneral.async = False
DocactivityGeneral.Load(server.MapPath(xmlfilename)) 
docactivityGeneral.setProperty "SelectionLanguage", "XPath"
set selectedactivityGeneralnode = docactivityGeneral.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedactivityGeneralnodes=docactivityGeneral.documentElement.selectNodes("/languages/language")
function getactivityGeneralLngStr(instring)
	temp = selectedactivityGeneralnode.selectSingleNode(instring).text
	getactivityGeneralLngStr = temp
end function
%>

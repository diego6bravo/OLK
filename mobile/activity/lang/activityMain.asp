<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "activityMain.xml"
set docactivityMain = server.CreateObject("MSXML2.DOMDocument")
docactivityMain.async = False
DocactivityMain.Load(server.MapPath(xmlfilename)) 
docactivityMain.setProperty "SelectionLanguage", "XPath"
set selectedactivityMainnode = docactivityMain.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedactivityMainnodes=docactivityMain.documentElement.selectNodes("/languages/language")
function getactivityMainLngStr(instring)
	temp = selectedactivityMainnode.selectSingleNode(instring).text
	getactivityMainLngStr = temp
end function
%>

<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "openActivities.xml"
set docopenActivities = server.CreateObject("MSXML2.DOMDocument")
docopenActivities.async = False
DocopenActivities.Load(server.MapPath(xmlfilename)) 
docopenActivities.setProperty "SelectionLanguage", "XPath"
set selectedopenActivitiesnode = docopenActivities.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedopenActivitiesnodes=docopenActivities.documentElement.selectNodes("/languages/language")
function getopenActivitiesLngStr(instring)
	temp = selectedopenActivitiesnode.selectSingleNode(instring).text
	getopenActivitiesLngStr = temp
end function
%>

<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "openActivitiesSearch.xml"
set docopenActivitiesSearch = server.CreateObject("MSXML2.DOMDocument")
docopenActivitiesSearch.async = False
DocopenActivitiesSearch.Load(server.MapPath(xmlfilename)) 
docopenActivitiesSearch.setProperty "SelectionLanguage", "XPath"
set selectedopenActivitiesSearchnode = docopenActivitiesSearch.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedopenActivitiesSearchnodes=docopenActivitiesSearch.documentElement.selectNodes("/languages/language")
function getopenActivitiesSearchLngStr(instring)
	temp = selectedopenActivitiesSearchnode.selectSingleNode(instring).text
	getopenActivitiesSearchLngStr = temp
end function
%>

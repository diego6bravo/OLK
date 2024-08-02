<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "openedSO.xml"
set docopenedSO = server.CreateObject("MSXML2.DOMDocument")
docopenedSO.async = False
DocopenedSO.Load(server.MapPath(xmlfilename)) 
docopenedSO.setProperty "SelectionLanguage", "XPath"
set selectedopenedSOnode = docopenedSO.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedopenedSOnodes=docopenedSO.documentElement.selectNodes("/languages/language")
function getopenedSOLngStr(instring)
	temp = selectedopenedSOnode.selectSingleNode(instring).text
	getopenedSOLngStr = temp
end function
%>

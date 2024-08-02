<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "ofertHistory.xml"
set docofertHistory = server.CreateObject("MSXML2.DOMDocument")
docofertHistory.async = False
DocofertHistory.Load(server.MapPath(xmlfilename)) 
docofertHistory.setProperty "SelectionLanguage", "XPath"
set selectedofertHistorynode = docofertHistory.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedofertHistorynodes=docofertHistory.documentElement.selectNodes("/languages/language")
function getofertHistoryLngStr(instring)
	temp = selectedofertHistorynode.selectSingleNode(instring).text
	getofertHistoryLngStr = temp
end function
%>

<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "taskMonitor.xml"
set doctaskMonitor = server.CreateObject("MSXML2.DOMDocument")
doctaskMonitor.async = False
DoctaskMonitor.Load(server.MapPath(xmlfilename)) 
doctaskMonitor.setProperty "SelectionLanguage", "XPath"
set selectedtaskMonitornode = doctaskMonitor.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedtaskMonitornodes=doctaskMonitor.documentElement.selectNodes("/languages/language")
function gettaskMonitorLngStr(instring)
	temp = selectedtaskMonitornode.selectSingleNode(instring).text
	gettaskMonitorLngStr = temp
end function
%>

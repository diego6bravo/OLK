<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminAlerts.xml"
set docadminAlerts = server.CreateObject("MSXML2.DOMDocument")
docadminAlerts.async = False
DocadminAlerts.Load(server.MapPath(xmlfilename)) 
docadminAlerts.setProperty "SelectionLanguage", "XPath"
set selectedadminAlertsnode = docadminAlerts.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminAlertsnodes=docadminAlerts.documentElement.selectNodes("/languages/language")
function getadminAlertsLngStr(instring)
	temp = selectedadminAlertsnode.selectSingleNode(instring).text
	getadminAlertsLngStr = temp
end function
%>

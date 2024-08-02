<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "messageAlert.xml"
set docmessageAlert = server.CreateObject("MSXML2.DOMDocument")
docmessageAlert.async = False
DocmessageAlert.Load(server.MapPath(xmlfilename)) 
docmessageAlert.setProperty "SelectionLanguage", "XPath"
set selectedmessageAlertnode = docmessageAlert.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedmessageAlertnodes=docmessageAlert.documentElement.selectNodes("/languages/language")
function getmessageAlertLngStr(instring)
	temp = selectedmessageAlertnode.selectSingleNode(instring).text
	getmessageAlertLngStr = temp
end function
%>

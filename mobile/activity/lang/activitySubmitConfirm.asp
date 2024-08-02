<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "activitySubmitConfirm.xml"
set docactivitySubmitConfirm = server.CreateObject("MSXML2.DOMDocument")
docactivitySubmitConfirm.async = False
DocactivitySubmitConfirm.Load(server.MapPath(xmlfilename)) 
docactivitySubmitConfirm.setProperty "SelectionLanguage", "XPath"
set selectedactivitySubmitConfirmnode = docactivitySubmitConfirm.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedactivitySubmitConfirmnodes=docactivitySubmitConfirm.documentElement.selectNodes("/languages/language")
function getactivitySubmitConfirmLngStr(instring)
	temp = selectedactivitySubmitConfirmnode.selectSingleNode(instring).text
	getactivitySubmitConfirmLngStr = temp
end function
%>

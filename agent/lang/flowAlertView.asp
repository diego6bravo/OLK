<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "flowAlertView.xml"
set docflowAlertView = server.CreateObject("MSXML2.DOMDocument")
docflowAlertView.async = False
DocflowAlertView.Load(server.MapPath(xmlfilename)) 
docflowAlertView.setProperty "SelectionLanguage", "XPath"
set selectedflowAlertViewnode = docflowAlertView.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedflowAlertViewnodes=docflowAlertView.documentElement.selectNodes("/languages/language")
function getflowAlertViewLngStr(instring)
	temp = selectedflowAlertViewnode.selectSingleNode(instring).text
	getflowAlertViewLngStr = temp
end function
%>

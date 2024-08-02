<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "flowAlert.xml"
set docflowAlert = server.CreateObject("MSXML2.DOMDocument")
docflowAlert.async = False
DocflowAlert.Load(server.MapPath(xmlfilename)) 
docflowAlert.setProperty "SelectionLanguage", "XPath"
set selectedflowAlertnode = docflowAlert.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedflowAlertnodes=docflowAlert.documentElement.selectNodes("/languages/language")
function getflowAlertLngStr(instring)
	temp = selectedflowAlertnode.selectSingleNode(instring).text
	getflowAlertLngStr = temp
end function
%>

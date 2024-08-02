<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "flowAlertDetails.xml"
set docflowAlertDetails = server.CreateObject("MSXML2.DOMDocument")
docflowAlertDetails.async = False
DocflowAlertDetails.Load(server.MapPath(xmlfilename)) 
docflowAlertDetails.setProperty "SelectionLanguage", "XPath"
set selectedflowAlertDetailsnode = docflowAlertDetails.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedflowAlertDetailsnodes=docflowAlertDetails.documentElement.selectNodes("/languages/language")
function getflowAlertDetailsLngStr(instring)
	temp = selectedflowAlertDetailsnode.selectSingleNode(instring).text
	getflowAlertDetailsLngStr = temp
end function
%>

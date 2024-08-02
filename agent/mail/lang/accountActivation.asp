<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "accountActivation.xml"
set docaccountActivation = server.CreateObject("MSXML2.DOMDocument")
docaccountActivation.async = False
DocaccountActivation.Load(server.MapPath(xmlfilename)) 
docaccountActivation.setProperty "SelectionLanguage", "XPath"
set selectedaccountActivationnode = docaccountActivation.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedaccountActivationnodes=docaccountActivation.documentElement.selectNodes("/languages/language")
function getaccountActivationLngStr(instring)
	temp = selectedaccountActivationnode.selectSingleNode(instring).text
	getaccountActivationLngStr = temp
end function
%>

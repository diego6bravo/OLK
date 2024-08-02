<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "confirmReload.xml"
set docconfirmReload = server.CreateObject("MSXML2.DOMDocument")
docconfirmReload.async = False
DocconfirmReload.Load(server.MapPath(xmlfilename)) 
docconfirmReload.setProperty "SelectionLanguage", "XPath"
set selectedconfirmReloadnode = docconfirmReload.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedconfirmReloadnodes=docconfirmReload.documentElement.selectNodes("/languages/language")
function getconfirmReloadLngStr(instring)
	temp = selectedconfirmReloadnode.selectSingleNode(instring).text
	getconfirmReloadLngStr = temp
end function
%>

<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminInvOpt.xml"
set docadminInvOpt = server.CreateObject("MSXML2.DOMDocument")
docadminInvOpt.async = False
DocadminInvOpt.Load(server.MapPath(xmlfilename)) 
docadminInvOpt.setProperty "SelectionLanguage", "XPath"
set selectedadminInvOptnode = docadminInvOpt.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminInvOptnodes=docadminInvOpt.documentElement.selectNodes("/languages/language")
function getadminInvOptLngStr(instring)
	temp = selectedadminInvOptnode.selectSingleNode(instring).text
	getadminInvOptLngStr = temp
end function
%>

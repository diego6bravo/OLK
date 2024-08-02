<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminInformer.xml"
set docadminInformer = server.CreateObject("MSXML2.DOMDocument")
docadminInformer.async = False
DocadminInformer.Load(server.MapPath(xmlfilename)) 
docadminInformer.setProperty "SelectionLanguage", "XPath"
set selectedadminInformernode = docadminInformer.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminInformernodes=docadminInformer.documentElement.selectNodes("/languages/language")
function getadminInformerLngStr(instring)
	temp = selectedadminInformernode.selectSingleNode(instring).text
	getadminInformerLngStr = temp
end function
%>

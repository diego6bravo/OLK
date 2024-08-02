<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminGeneral.xml"
set docadminGeneral = server.CreateObject("MSXML2.DOMDocument")
docadminGeneral.async = False
DocadminGeneral.Load(server.MapPath(xmlfilename)) 
docadminGeneral.setProperty "SelectionLanguage", "XPath"
set selectedadminGeneralnode = docadminGeneral.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminGeneralnodes=docadminGeneral.documentElement.selectNodes("/languages/language")
function getadminGeneralLngStr(instring)
	temp = selectedadminGeneralnode.selectSingleNode(instring).text
	getadminGeneralLngStr = temp
end function
%>

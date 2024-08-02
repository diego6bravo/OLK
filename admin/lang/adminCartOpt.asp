<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminCartOpt.xml"
set docadminCartOpt = server.CreateObject("MSXML2.DOMDocument")
docadminCartOpt.async = False
DocadminCartOpt.Load(server.MapPath(xmlfilename)) 
docadminCartOpt.setProperty "SelectionLanguage", "XPath"
set selectedadminCartOptnode = docadminCartOpt.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminCartOptnodes=docadminCartOpt.documentElement.selectNodes("/languages/language")
function getadminCartOptLngStr(instring)
	temp = selectedadminCartOptnode.selectSingleNode(instring).text
	getadminCartOptLngStr = temp
end function
%>

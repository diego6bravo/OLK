<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminRepNew.xml"
set docadminRepNew = server.CreateObject("MSXML2.DOMDocument")
docadminRepNew.async = False
DocadminRepNew.Load(server.MapPath(xmlfilename)) 
docadminRepNew.setProperty "SelectionLanguage", "XPath"
set selectedadminRepNewnode = docadminRepNew.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminRepNewnodes=docadminRepNew.documentElement.selectNodes("/languages/language")
function getadminRepNewLngStr(instring)
	temp = selectedadminRepNewnode.selectSingleNode(instring).text
	getadminRepNewLngStr = temp
end function
%>

<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminReps.xml"
set docadminReps = server.CreateObject("MSXML2.DOMDocument")
docadminReps.async = False
DocadminReps.Load(server.MapPath(xmlfilename)) 
docadminReps.setProperty "SelectionLanguage", "XPath"
set selectedadminRepsnode = docadminReps.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminRepsnodes=docadminReps.documentElement.selectNodes("/languages/language")
function getadminRepsLngStr(instring)
	temp = selectedadminRepsnode.selectSingleNode(instring).text
	getadminRepsLngStr = temp
end function
%>

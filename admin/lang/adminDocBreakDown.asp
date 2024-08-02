<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminDocBreakDown.xml"
set docadminDocBreakDown = server.CreateObject("MSXML2.DOMDocument")
docadminDocBreakDown.async = False
DocadminDocBreakDown.Load(server.MapPath(xmlfilename)) 
docadminDocBreakDown.setProperty "SelectionLanguage", "XPath"
set selectedadminDocBreakDownnode = docadminDocBreakDown.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminDocBreakDownnodes=docadminDocBreakDown.documentElement.selectNodes("/languages/language")
function getadminDocBreakDownLngStr(instring)
	temp = selectedadminDocBreakDownnode.selectSingleNode(instring).text
	getadminDocBreakDownLngStr = temp
end function
%>

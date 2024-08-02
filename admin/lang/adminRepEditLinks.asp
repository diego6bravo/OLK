<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminRepEditLinks.xml"
set docadminRepEditLinks = server.CreateObject("MSXML2.DOMDocument")
docadminRepEditLinks.async = False
DocadminRepEditLinks.Load(server.MapPath(xmlfilename)) 
docadminRepEditLinks.setProperty "SelectionLanguage", "XPath"
set selectedadminRepEditLinksnode = docadminRepEditLinks.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminRepEditLinksnodes=docadminRepEditLinks.documentElement.selectNodes("/languages/language")
function getadminRepEditLinksLngStr(instring)
	temp = selectedadminRepEditLinksnode.selectSingleNode(instring).text
	getadminRepEditLinksLngStr = temp
end function
%>

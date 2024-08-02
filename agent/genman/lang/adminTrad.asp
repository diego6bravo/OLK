<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminTrad.xml"
set docadminTrad = server.CreateObject("MSXML2.DOMDocument")
docadminTrad.async = False
DocadminTrad.Load(server.MapPath(xmlfilename)) 
docadminTrad.setProperty "SelectionLanguage", "XPath"
set selectedadminTradnode = docadminTrad.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminTradnodes=docadminTrad.documentElement.selectNodes("/languages/language")
function getadminTradLngStr(instring)
	temp = selectedadminTradnode.selectSingleNode(instring).text
	getadminTradLngStr = temp
end function
%>

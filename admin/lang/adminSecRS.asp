<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminSecRS.xml"
set docadminSecRS = server.CreateObject("MSXML2.DOMDocument")
docadminSecRS.async = False
DocadminSecRS.Load(server.MapPath(xmlfilename)) 
docadminSecRS.setProperty "SelectionLanguage", "XPath"
set selectedadminSecRSnode = docadminSecRS.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminSecRSnodes=docadminSecRS.documentElement.selectNodes("/languages/language")
function getadminSecRSLngStr(instring)
	temp = selectedadminSecRSnode.selectSingleNode(instring).text
	getadminSecRSLngStr = temp
end function
%>

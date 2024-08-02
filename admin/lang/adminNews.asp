<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminNews.xml"
set docadminNews = server.CreateObject("MSXML2.DOMDocument")
docadminNews.async = False
DocadminNews.Load(server.MapPath(xmlfilename)) 
docadminNews.setProperty "SelectionLanguage", "XPath"
set selectedadminNewsnode = docadminNews.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminNewsnodes=docadminNews.documentElement.selectNodes("/languages/language")
function getadminNewsLngStr(instring)
	temp = selectedadminNewsnode.selectSingleNode(instring).text
	getadminNewsLngStr = temp
end function
%>

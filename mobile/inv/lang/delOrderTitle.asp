<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "delOrderTitle.xml"
set docdelOrderTitle = server.CreateObject("MSXML2.DOMDocument")
docdelOrderTitle.async = False
DocdelOrderTitle.Load(server.MapPath(xmlfilename)) 
docdelOrderTitle.setProperty "SelectionLanguage", "XPath"
set selecteddelOrderTitlenode = docdelOrderTitle.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selecteddelOrderTitlenodes=docdelOrderTitle.documentElement.selectNodes("/languages/language")
function getdelOrderTitleLngStr(instring)
	temp = selecteddelOrderTitlenode.selectSingleNode(instring).text
	getdelOrderTitleLngStr = temp
end function
%>

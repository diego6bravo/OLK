<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "docdel.xml"
set docdocdel = server.CreateObject("MSXML2.DOMDocument")
docdocdel.async = False
Docdocdel.Load(server.MapPath(xmlfilename)) 
docdocdel.setProperty "SelectionLanguage", "XPath"
set selecteddocdelnode = docdocdel.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selecteddocdelnodes=docdocdel.documentElement.selectNodes("/languages/language")
function getdocdelLngStr(instring)
	temp = selecteddocdelnode.selectSingleNode(instring).text
	getdocdelLngStr = temp
end function
%>

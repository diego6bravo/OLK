<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "openDocs.xml"
set docopenDocs = server.CreateObject("MSXML2.DOMDocument")
docopenDocs.async = False
DocopenDocs.Load(server.MapPath(xmlfilename)) 
docopenDocs.setProperty "SelectionLanguage", "XPath"
set selectedopenDocsnode = docopenDocs.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedopenDocsnodes=docopenDocs.documentElement.selectNodes("/languages/language")
function getopenDocsLngStr(instring)
	temp = selectedopenDocsnode.selectSingleNode(instring).text
	getopenDocsLngStr = temp
end function
%>

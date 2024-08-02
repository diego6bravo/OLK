<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "reload.xml"
set docreload = server.CreateObject("MSXML2.DOMDocument")
docreload.async = False
Docreload.Load(server.MapPath(xmlfilename)) 
docreload.setProperty "SelectionLanguage", "XPath"
set selectedreloadnode = docreload.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedreloadnodes=docreload.documentElement.selectNodes("/languages/language")
function getreloadLngStr(instring)
	temp = selectedreloadnode.selectSingleNode(instring).text
	getreloadLngStr = temp
end function
%>

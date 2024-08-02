<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "listpen.xml"
set doclistpen = server.CreateObject("MSXML2.DOMDocument")
doclistpen.async = False
Doclistpen.Load(server.MapPath(xmlfilename)) 
doclistpen.setProperty "SelectionLanguage", "XPath"
set selectedlistpennode = doclistpen.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedlistpennodes=doclistpen.documentElement.selectNodes("/languages/language")
function getlistpenLngStr(instring)
	temp = selectedlistpennode.selectSingleNode(instring).text
	getlistpenLngStr = temp
end function
%>

<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "newactgo.xml"
set docnewactgo = server.CreateObject("MSXML2.DOMDocument")
docnewactgo.async = False
Docnewactgo.Load(server.MapPath(xmlfilename)) 
docnewactgo.setProperty "SelectionLanguage", "XPath"
set selectednewactgonode = docnewactgo.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectednewactgonodes=docnewactgo.documentElement.selectNodes("/languages/language")
function getnewactgoLngStr(instring)
	temp = selectednewactgonode.selectSingleNode(instring).text
	getnewactgoLngStr = temp
end function
%>

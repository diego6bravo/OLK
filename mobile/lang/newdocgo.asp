<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "newdocgo.xml"
set docnewdocgo = server.CreateObject("MSXML2.DOMDocument")
docnewdocgo.async = False
Docnewdocgo.Load(server.MapPath(xmlfilename)) 
docnewdocgo.setProperty "SelectionLanguage", "XPath"
set selectednewdocgonode = docnewdocgo.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectednewdocgonodes=docnewdocgo.documentElement.selectNodes("/languages/language")
function getnewdocgoLngStr(instring)
	temp = selectednewdocgonode.selectSingleNode(instring).text
	getnewdocgoLngStr = temp
end function
%>

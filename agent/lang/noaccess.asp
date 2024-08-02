<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "noaccess.xml"
set docnoaccess = server.CreateObject("MSXML2.DOMDocument")
docnoaccess.async = False
Docnoaccess.Load(server.MapPath(xmlfilename)) 
docnoaccess.setProperty "SelectionLanguage", "XPath"
set selectednoaccessnode = docnoaccess.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectednoaccessnodes=docnoaccess.documentElement.selectNodes("/languages/language")
function getnoaccessLngStr(instring)
	temp = selectednoaccessnode.selectSingleNode(instring).text
	getnoaccessLngStr = temp
end function
%>

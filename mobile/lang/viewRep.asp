<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "viewRep.xml"
set docviewRep = server.CreateObject("MSXML2.DOMDocument")
docviewRep.async = False
DocviewRep.Load(server.MapPath(xmlfilename)) 
docviewRep.setProperty "SelectionLanguage", "XPath"
set selectedviewRepnode = docviewRep.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedviewRepnodes=docviewRep.documentElement.selectNodes("/languages/language")
function getviewRepLngStr(instring)
	temp = selectedviewRepnode.selectSingleNode(instring).text
	getviewRepLngStr = temp
end function
%>

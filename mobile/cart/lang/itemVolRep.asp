<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "itemVolRep.xml"
set docitemVolRep = server.CreateObject("MSXML2.DOMDocument")
docitemVolRep.async = False
DocitemVolRep.Load(server.MapPath(xmlfilename)) 
docitemVolRep.setProperty "SelectionLanguage", "XPath"
set selecteditemVolRepnode = docitemVolRep.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selecteditemVolRepnodes=docitemVolRep.documentElement.selectNodes("/languages/language")
function getitemVolRepLngStr(instring)
	temp = selecteditemVolRepnode.selectSingleNode(instring).text
	getitemVolRepLngStr = temp
end function
%>

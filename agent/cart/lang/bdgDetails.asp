<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "bdgDetails.xml"
set docbdgDetails = server.CreateObject("MSXML2.DOMDocument")
docbdgDetails.async = False
DocbdgDetails.Load(server.MapPath(xmlfilename)) 
docbdgDetails.setProperty "SelectionLanguage", "XPath"
set selectedbdgDetailsnode = docbdgDetails.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedbdgDetailsnodes=docbdgDetails.documentElement.selectNodes("/languages/language")
function getbdgDetailsLngStr(instring)
	temp = selectedbdgDetailsnode.selectSingleNode(instring).text
	getbdgDetailsLngStr = temp
end function
%>

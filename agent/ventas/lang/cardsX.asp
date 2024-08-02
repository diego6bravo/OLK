<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "cardsX.xml"
set doccardsX = server.CreateObject("MSXML2.DOMDocument")
doccardsX.async = False
DoccardsX.Load(server.MapPath(xmlfilename)) 
doccardsX.setProperty "SelectionLanguage", "XPath"
set selectedcardsXnode = doccardsX.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedcardsXnodes=doccardsX.documentElement.selectNodes("/languages/language")
function getcardsXLngStr(instring)
	temp = selectedcardsXnode.selectSingleNode(instring).text
	getcardsXLngStr = temp
end function
%>

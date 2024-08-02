<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "ventasX.xml"
set docventasX = server.CreateObject("MSXML2.DOMDocument")
docventasX.async = False
DocventasX.Load(server.MapPath(xmlfilename)) 
docventasX.setProperty "SelectionLanguage", "XPath"
set selectedventasXnode = docventasX.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedventasXnodes=docventasX.documentElement.selectNodes("/languages/language")
function getventasXLngStr(instring)
	temp = selectedventasXnode.selectSingleNode(instring).text
	getventasXLngStr = temp
end function
%>

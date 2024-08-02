<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "itemsX.xml"
set docitemsX = server.CreateObject("MSXML2.DOMDocument")
docitemsX.async = False
DocitemsX.Load(server.MapPath(xmlfilename)) 
docitemsX.setProperty "SelectionLanguage", "XPath"
set selecteditemsXnode = docitemsX.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selecteditemsXnodes=docitemsX.documentElement.selectNodes("/languages/language")
function getitemsXLngStr(instring)
	temp = selecteditemsXnode.selectSingleNode(instring).text
	getitemsXLngStr = temp
end function
%>

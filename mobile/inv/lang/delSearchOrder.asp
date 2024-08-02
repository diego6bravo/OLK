<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "delSearchOrder.xml"
set docdelSearchOrder = server.CreateObject("MSXML2.DOMDocument")
docdelSearchOrder.async = False
DocdelSearchOrder.Load(server.MapPath(xmlfilename)) 
docdelSearchOrder.setProperty "SelectionLanguage", "XPath"
set selecteddelSearchOrdernode = docdelSearchOrder.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selecteddelSearchOrdernodes=docdelSearchOrder.documentElement.selectNodes("/languages/language")
function getdelSearchOrderLngStr(instring)
	temp = selecteddelSearchOrdernode.selectSingleNode(instring).text
	getdelSearchOrderLngStr = temp
end function
%>
